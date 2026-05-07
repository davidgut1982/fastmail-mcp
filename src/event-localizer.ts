import { DateTime } from 'luxon';
import { resolveDefaultTimezone } from './timezone.js';

/**
 * Why: The CalDAV server returns event timestamps as UTC instants
 * (e.g. `2026-05-07T13:00:00Z`). When the agent reads those back, it parses
 * "13:00 UTC" naively and decides the event is in the AFTERNOON, so a query
 * for "morning" filters it out client-side — even though 13:00Z = 09:00 EDT,
 * which IS morning in the user's actual timezone. Surfacing the local-zone
 * representation (with explicit offset) at the MCP boundary removes that
 * misread: the agent sees `2026-05-07T09:00:00-04:00` and correctly classifies
 * it as morning. We also keep the original UTC instant as `startUtc`/`endUtc`
 * so any downstream call that wants to round-trip the value back to
 * `update_calendar_event` can do so without shifting the moment.
 *
 * What: `localizeEventTimes` takes an event with UTC-formatted `start`/`end`
 * strings (or all-day date-only strings) and returns a new event whose
 * `start`/`end` are the same instant expressed in `tz` with an explicit
 * offset, plus `startUtc`/`endUtc` preserving the original UTC verbatim and
 * a `timezone` field naming the target zone. All-day inputs (date-only,
 * no `T`) pass through unchanged — they have no time-of-day to localize and
 * mangling them into datetimes would corrupt the all-day semantics.
 *
 * Test: With tz="America/New_York", an event with start="2026-05-07T13:00:00Z"
 * is returned with start="2026-05-07T09:00:00.000-04:00",
 * startUtc="2026-05-07T13:00:00.000Z", timezone="America/New_York".
 */

/**
 * Why: All-day events are signaled by a date-only `DTSTART` (no time portion).
 * In our parsed CalendarEvent shape `formatICalDate` returns those as
 * `YYYY-MM-DD` (no `T`). We must NOT shove a 00:00:00 time onto them — that
 * would turn an all-day birthday into a midnight-local datetime, which is a
 * different concept. Keep date-only strings as-is.
 *
 * What: Returns true iff the input is a `YYYY-MM-DD` string with no `T`.
 *
 * Test: isAllDayDateString('2026-05-07') === true,
 * isAllDayDateString('2026-05-07T13:00:00Z') === false,
 * isAllDayDateString(undefined) === false.
 */
export function isAllDayDateString(value: string | undefined): boolean {
  if (typeof value !== 'string') return false;
  return /^\d{4}-\d{2}-\d{2}$/.test(value.trim());
}

/**
 * Convert a single UTC ISO string to the target zone with explicit offset.
 *
 * Why: Centralizes the luxon conversion so every callsite handles invalid
 * inputs identically (return undefined rather than throwing — a malformed
 * server response should not blow up the whole list response).
 *
 * What: Parses `utcIso` as UTC, rezones to `tz`, and emits an ISO string
 * with explicit offset (e.g. `2026-05-07T09:00:00.000-04:00`). Returns
 * `undefined` if the input is not parseable.
 *
 * Test: convertUtcToZone('2026-05-07T13:00:00Z', 'America/New_York') ===
 * '2026-05-07T09:00:00.000-04:00' (EDT in May = UTC-4).
 */
export function convertUtcToZone(utcIso: string, tz: string): string | undefined {
  // Accept inputs both with and without the `Z` suffix — formatICalDate emits
  // `2026-05-07T13:00:00Z` for UTC-stamped DTSTARTs but `2026-05-07T08:30:00`
  // for floating-time DTSTARTs. The latter shouldn't happen on the read path
  // for events the user cares about, but be defensive.
  const dt = DateTime.fromISO(utcIso, { zone: 'utc' });
  if (!dt.isValid) return undefined;
  const zoned = dt.setZone(tz);
  if (!zoned.isValid) return undefined;
  // includeOffset=true (default) emits the explicit `-04:00` style suffix so
  // the agent can't misread the value as UTC. Keep millisecond precision so
  // round-tripping through Date.parse is lossless.
  return zoned.toISO() ?? undefined;
}

/**
 * Localized form of CalendarEvent. Extends the base shape with the new
 * timezone-aware fields. Marked optional because callers that consume the
 * un-localized form (legacy tests, internal call paths) should still work.
 */
export interface LocalizedCalendarEventFields {
  startUtc?: string;
  endUtc?: string;
  timezone?: string;
}

/**
 * Why: This is the function called at every MCP-boundary read path so that
 * the JSON payload the agent sees encodes start/end in the user's local
 * zone with an explicit offset. The agent's afternoon/morning heuristic
 * then matches the user's mental model (their wall clock), eliminating the
 * 09:00 EDT → 13:00Z → "afternoon, drop it" misread.
 *
 * What: Returns a new event object with:
 *   - `start` / `end`: rezoned to `tz` with explicit offset (e.g. `-04:00`)
 *   - `startUtc` / `endUtc`: ORIGINAL UTC values verbatim (so round-tripping
 *     through update_calendar_event is byte-exact and won't shift the instant)
 *   - `timezone`: the IANA zone name, so the agent has explicit signal
 *     that the times are local
 * All-day events (date-only `start`/`end`, no `T`) pass through with their
 * original date string preserved; only `timezone` is added so the agent has
 * the consistent shape signal.
 *
 * Test: With tz="America/New_York":
 * - timed event: start="2026-05-07T13:00:00Z" → start="2026-05-07T09:00:00.000-04:00",
 *   startUtc="2026-05-07T13:00:00.000Z", timezone="America/New_York"
 * - all-day event: start="2026-05-07" → start="2026-05-07" unchanged,
 *   timezone="America/New_York"
 *
 * @param event Raw CalendarEvent shape with UTC-formatted (or date-only) start/end.
 * @param tz Target IANA zone. Defaults to resolveDefaultTimezone() so callers
 *   in production don't need to pass it; tests pass an explicit zone.
 */
export function localizeEventTimes<T extends { start?: string; end?: string }>(
  event: T,
  tz?: string,
): T & LocalizedCalendarEventFields {
  const zone = tz || resolveDefaultTimezone();

  // Defensive: never mutate the input. The caller may still hold the
  // un-localized object (e.g. tests inspect the parsed-but-pre-localized
  // form), so produce a new object.
  const out: T & LocalizedCalendarEventFields = { ...event, timezone: zone };

  // Handle start. All-day → leave date string alone; timed → convert + preserve UTC.
  if (typeof event.start === 'string' && event.start.length > 0) {
    if (isAllDayDateString(event.start)) {
      // Date-only: pass through verbatim. Do NOT set startUtc — the value
      // is not a UTC instant, it's a calendar date, and emitting a fake
      // "midnight-UTC" startUtc would mislead consumers about the event's
      // semantics (all-day vs. timed).
      out.start = event.start;
    } else {
      const localized = convertUtcToZone(event.start, zone);
      if (localized) {
        out.start = localized;
        // Preserve the ORIGINAL string verbatim. Per spec: "must contain the
        // original CalDAV UTC value verbatim — not a re-derived value, in
        // case rounding/parsing shifts something".
        out.startUtc = event.start;
      }
      // If conversion fails (malformed input), leave start as-is and don't
      // set startUtc. Better to surface bad data than to drop the field.
    }
  }

  if (typeof event.end === 'string' && event.end.length > 0) {
    if (isAllDayDateString(event.end)) {
      out.end = event.end;
    } else {
      const localized = convertUtcToZone(event.end, zone);
      if (localized) {
        out.end = localized;
        out.endUtc = event.end;
      }
    }
  }

  return out;
}
