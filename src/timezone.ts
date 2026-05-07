import { DateTime, IANAZone } from 'luxon';

/**
 * Why: When the agent passes a naive datetime like `2026-05-07T00:00:00`
 * (no `Z`, no offset), interpreting it as UTC silently shifts the user's
 * "Thursday morning" 4-5 hours and excludes their actual morning meetings.
 * Centralizing the user's local timezone — resolved once at startup, with
 * an explicit override knob — lets us honor an explicit offset when present
 * but treat naive inputs as wall-clock time in the user's local zone.
 *
 * What: Resolves the configured default timezone with a 4-step fallback:
 *   1. FASTMAIL_TIMEZONE env var (explicit override)
 *   2. process.env.TZ (POSIX convention, also set by systemd Environment=TZ=...)
 *   3. System default via Intl.DateTimeFormat().resolvedOptions().timeZone
 *   4. Hard fallback: 'America/New_York' (the user we ship to)
 * Each candidate is validated with `IANAZone.isValidZone` before being
 * accepted; invalid candidates are skipped with a warning so a typo in env
 * doesn't silently re-introduce the UTC bug.
 *
 * Test: Unset FASTMAIL_TIMEZONE and TZ, mock Intl to return "Europe/London",
 * assert resolveDefaultTimezone() returns "Europe/London". Set
 * FASTMAIL_TIMEZONE=America/New_York, assert it returns "America/New_York".
 * Set FASTMAIL_TIMEZONE to garbage, assert it falls through to TZ.
 */
function isValidIanaTz(tz: string): boolean {
  if (!tz || typeof tz !== 'string') return false;
  // luxon's IANAZone validation is the same lookup the runtime uses for
  // conversions, so a passing check here means downstream conversions won't
  // throw "Invalid time zone specified".
  if (!IANAZone.isValidZone(tz)) return false;
  // Belt-and-suspenders: also verify Intl accepts it. Some Node builds have
  // ICU stripped down and reject zones that luxon's static list accepts.
  try {
    new Intl.DateTimeFormat(undefined, { timeZone: tz });
    return true;
  } catch {
    return false;
  }
}

let cachedDefaultTimezone: string | null = null;

export function resolveDefaultTimezone(): string {
  if (cachedDefaultTimezone) return cachedDefaultTimezone;

  const candidates: Array<{ source: string; value: string | undefined }> = [
    { source: 'FASTMAIL_TIMEZONE', value: process.env.FASTMAIL_TIMEZONE },
    { source: 'TZ', value: process.env.TZ },
    {
      source: 'system (Intl)',
      value: (() => {
        try {
          return Intl.DateTimeFormat().resolvedOptions().timeZone;
        } catch {
          return undefined;
        }
      })(),
    },
    { source: 'fallback', value: 'America/New_York' },
  ];

  for (const c of candidates) {
    if (!c.value) continue;
    if (isValidIanaTz(c.value)) {
      cachedDefaultTimezone = c.value;
      console.error(
        `[fastmail-mcp] Resolved default timezone: ${c.value} (source: ${c.source})`,
      );
      return cachedDefaultTimezone;
    }
    console.error(
      `[fastmail-mcp] Ignoring invalid IANA timezone "${c.value}" from ${c.source}`,
    );
  }

  // The hard fallback above is itself a valid IANA zone, so this should be
  // unreachable. Keep a defensive bail to satisfy the type system.
  cachedDefaultTimezone = 'America/New_York';
  return cachedDefaultTimezone;
}

/**
 * Why: We need to distinguish "agent passed a UTC instant" from "agent passed
 * a wall-clock time the user meant in their local zone". The presence of a
 * `Z` suffix or `[+-]HH:MM` (or `[+-]HHMM`) offset means the caller has
 * already committed to a UTC instant; without one, the value is a naive
 * civil time that should be interpreted in the user's configured zone.
 *
 * What: Returns true iff the input string ends with `Z`, or contains an
 * explicit offset like `+05:00`, `-0500`, `+05`, `-05` near the end. Matches
 * conservatively — only after a `T...HH:MM[:SS][.fff]` time portion, to
 * avoid mistaking a date hyphen (`2026-05-07`) for a negative offset.
 *
 * Test: Assert hasExplicitOffset('2026-05-07T00:00:00Z') === true,
 * hasExplicitOffset('2026-05-07T00:00:00-05:00') === true,
 * hasExplicitOffset('2026-05-07T00:00:00+0500') === true,
 * hasExplicitOffset('2026-05-07T00:00:00') === false,
 * hasExplicitOffset('2026-05-07') === false.
 */
export function hasExplicitOffset(input: string): boolean {
  if (typeof input !== 'string') return false;
  const trimmed = input.trim();
  if (trimmed.endsWith('Z') || trimmed.endsWith('z')) return true;
  // Look for an offset attached to the time portion: T...[+-]HH(:?MM)?
  // Anchored at end of string to avoid catching the date's hyphens.
  return /T\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?[+-]\d{2}:?\d{0,2}$/.test(trimmed);
}

/**
 * Why: This is the hinge of the whole TZ-aware refactor. Every datetime
 * input crossing the MCP boundary — query window edges in
 * get_calendar_events, event start/end in create_calendar_event /
 * update_calendar_event — must flow through here so that "2026-05-07T00:00:00"
 * means 04:00Z (in EDT) and not 00:00Z.
 *
 * What: Returns a UTC ISO 8601 string (always with `Z` suffix) suitable for
 * downstream CalDAV formatting. If the input has an explicit offset/Z, it's
 * parsed as-is. If it's naive (date-only or date+time without offset), it's
 * interpreted as a wall-clock time in the configured default timezone (or
 * the explicit `tz` arg, when provided for testing).
 *
 * Test: With tz="America/New_York" (UTC-4 in May), assert
 * resolveDateInput('2026-05-07T00:00:00') === '2026-05-07T04:00:00.000Z',
 * resolveDateInput('2026-05-07T00:00:00Z') === '2026-05-07T00:00:00.000Z',
 * resolveDateInput('2026-05-07T00:00:00-05:00') === '2026-05-07T05:00:00.000Z'.
 */
export function resolveDateInput(input: string, tz?: string): string {
  if (typeof input !== 'string' || input.trim().length === 0) {
    throw new Error(`Invalid datetime input: ${JSON.stringify(input)}`);
  }
  const trimmed = input.trim();
  const zone = tz || resolveDefaultTimezone();

  let dt: DateTime;
  if (hasExplicitOffset(trimmed)) {
    // Parse respecting the embedded offset; setZone('utc') normalizes for output.
    dt = DateTime.fromISO(trimmed, { setZone: true });
    if (!dt.isValid) {
      throw new Error(
        `Invalid datetime input "${input}": ${dt.invalidReason} ${dt.invalidExplanation || ''}`.trim(),
      );
    }
    return dt.toUTC().toISO() as string;
  }

  // Naive input — interpret in the configured zone.
  // Date-only ("2026-05-07") becomes 00:00:00 in the zone, which is the most
  // common-sense interpretation for whole-day boundaries.
  dt = DateTime.fromISO(trimmed, { zone });
  if (!dt.isValid) {
    throw new Error(
      `Invalid datetime input "${input}": ${dt.invalidReason} ${dt.invalidExplanation || ''}`.trim(),
    );
  }
  return dt.toUTC().toISO() as string;
}

/**
 * Reset cached default timezone. Test-only — production paths should never
 * call this, since the resolved zone is meant to be stable for the lifetime
 * of the process.
 */
export function __resetTimezoneCacheForTests(): void {
  cachedDefaultTimezone = null;
}
