import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';
import { resolveDateInput } from './timezone.js';
import { localizeEventTimes } from './event-localizer.js';

/**
 * Convert ISO-8601 datetime (with or without timezone offset) to RFC 5545 UTC form.
 * Output format: YYYYMMDDTHHMMSSZ
 *
 * The previous implementation `event.start.replace(/[-:]/g, '')` destroyed timezone
 * offsets by stripping `:` from `+05:00`, producing invalid CalDAV strings that the
 * server silently rejected (or accepted as garbage). This helper goes through Date
 * to normalize to UTC, then formats per RFC 5545.
 */
export function toCalDateUTC(iso: string): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) {
    throw new Error(`Invalid datetime: ${iso}`);
  }
  const y = d.getUTCFullYear().toString().padStart(4, '0');
  const m = (d.getUTCMonth() + 1).toString().padStart(2, '0');
  const day = d.getUTCDate().toString().padStart(2, '0');
  const hh = d.getUTCHours().toString().padStart(2, '0');
  const mm = d.getUTCMinutes().toString().padStart(2, '0');
  const ss = d.getUTCSeconds().toString().padStart(2, '0');
  return `${y}${m}${day}T${hh}${mm}${ss}Z`;
}

export interface CalDAVConfig {
  username: string;
  password: string;
  serverUrl?: string;
}

export interface CalendarInfo {
  id: string;
  displayName: string;
  url: string;
  description?: string;
  color?: string;
}

/**
 * A parsed VALARM alarm (RFC 5545 §3.6.6).
 * `trigger` is the raw trigger value, e.g. `-PT15M` (15 min before),
 * `-P1D` (1 day before), or an absolute `YYYYMMDDTHHMMSSZ`.
 */
export interface ParsedAlarm {
  action: string;
  trigger: string;
  description?: string;
}

/**
 * A parsed RRULE recurrence rule (RFC 5545 §3.3.10).
 * Only the common parts are surfaced; the raw value is in `raw`.
 */
export interface ParsedRecurrence {
  freq?: string;
  interval?: number;
  count?: number;
  until?: string;
  byDay?: string[];
  byMonthDay?: number[];
  raw: string;
}

export interface CalendarEvent {
  id: string;
  url: string;
  title: string;
  description?: string;
  /**
   * Event start. On the read path (after localizeEventTimes), this is an
   * ISO 8601 string in the user's configured timezone with explicit offset
   * (e.g. `2026-05-07T09:00:00.000-04:00`), OR a `YYYY-MM-DD` date for
   * all-day events. The original UTC instant is preserved in `startUtc`.
   */
  start?: string;
  /** Event end. Same semantics as `start`; UTC original in `endUtc`. */
  end?: string;
  location?: string;
  /**
   * Original UTC instant for `start`, exactly as the CalDAV server returned
   * it (e.g. `2026-05-07T13:00:00Z`). Use this when round-tripping back to
   * `update_calendar_event` so rounding/parsing can't shift the moment.
   * Absent for all-day events (date-only `start`).
   */
  startUtc?: string;
  /** Original UTC instant for `end`. Same semantics as `startUtc`. */
  endUtc?: string;
  /**
   * IANA zone (e.g. `America/New_York`) the `start`/`end` strings are
   * expressed in. Explicit signal to consumers that times are LOCAL, not
   * UTC — so reasoning about morning/afternoon/evening uses the user's
   * wall clock, not UTC.
   */
  timezone?: string;
  /**
   * Event attendees parsed from ICS ATTENDEE lines. Each entry carries the
   * mailto address, optional CN display name, and whether RSVP was requested.
   * Absent when the VEVENT has no ATTENDEE lines (non-scheduling event).
   */
  attendees?: Array<{ email: string; name?: string; rsvp?: boolean }>;
  /** Whether this is an all-day event (DTSTART;VALUE=DATE). */
  allDay?: boolean;
  /** Event status (RFC 5545 §3.8.1.11). */
  status?: string;
  /** Free/busy transparency (RFC 5545 §3.8.2.7). OPAQUE=busy, TRANSPARENT=free. */
  transparency?: string;
  /** Event priority 0-9 (RFC 5545 §3.8.1.9). 0=undefined, 1=highest, 9=lowest. */
  priority?: number;
  /** Privacy class (RFC 5545 §3.8.1.3). PUBLIC, PRIVATE, or CONFIDENTIAL. */
  classification?: string;
  /** Categories/tags (RFC 5545 §3.8.1.2). */
  categories?: string[];
  /** Alarms parsed from VALARM blocks (RFC 5545 §3.6.6). */
  alarms?: ParsedAlarm[];
  /** Recurrence rule (RFC 5545 §3.3.10). Only present on the master VEVENT. */
  recurrence?: ParsedRecurrence;
  /** Organizer parsed from ORGANIZER line (RFC 5545 §3.8.4.3). */
  organizer?: { email: string; name?: string };
  /** Revision sequence number (RFC 5545 §3.8.7.4). */
  sequence?: number;
}

/**
 * Parse ATTENDEE lines out of a VEVENT/ICS block.
 *
 * Why: get_calendar_event / list_calendar_events previously dropped attendee
 * information entirely, so the agent could never tell who was invited to a
 * meeting. This extracts the mailto address, CN display name, and RSVP flag.
 * What: Returns an array of {email, name?, rsvp?} or undefined when none found.
 * Test: Feed an ICS containing `ATTENDEE;CN="John Doe";RSVP=TRUE:mailto:john@example.com`
 *       and assert one entry {email:'john@example.com', name:'John Doe', rsvp:true}.
 */
export function parseAttendees(
  icsText: string,
): Array<{ email: string; name?: string; rsvp?: boolean }> | undefined {
  const attendeeLines = icsText.match(/^ATTENDEE[^:]*:mailto:(.+)$/gim) || [];
  if (attendeeLines.length === 0) return undefined;
  return attendeeLines.map(line => {
    const email = line.split(':mailto:')[1]?.trim() || '';
    const cnMatch = line.match(/CN="?([^";:]+)"?/i);
    const name = cnMatch ? cnMatch[1].trim() : undefined;
    const rsvp = /RSVP=TRUE/i.test(line);
    return { email, name, rsvp };
  });
}

/**
 * Parse all VALARM blocks from a VEVENT (RFC 5545 §3.6.6).
 *
 * Why: Previously, alarms were silently dropped on read. The agent couldn't
 * see existing alarms, and the original use case (add reminders to events)
 * was impossible because the data was invisible.
 *
 * What: Extracts ACTION, TRIGGER, and DESCRIPTION from each VALARM block
 * within the VEVENT. Returns undefined when no VALARM is present.
 *
 * Test: Feed ICS with `BEGIN:VALARM\r\nTRIGGER:-PT15M\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nEND:VALARM`
 *       and assert [{action:'DISPLAY', trigger:'-PT15M', description:'Reminder'}].
 */
export function parseAlarms(vevent: string): ParsedAlarm[] | undefined {
  const alarmRe = /BEGIN:VALARM[\s\S]*?END:VALARM/gi;
  const alarms: ParsedAlarm[] = [];
  let m: RegExpExecArray | null;
  while ((m = alarmRe.exec(vevent)) !== null) {
    const block = m[0];
    const action = parseICalValue(block, 'ACTION') || 'DISPLAY';
    const trigger = parseICalValue(block, 'TRIGGER') || '';
    const description = parseICalValue(block, 'DESCRIPTION');
    alarms.push({
      action: action.toUpperCase(),
      trigger,
      description: description || undefined,
    });
  }
  return alarms.length > 0 ? alarms : undefined;
}

/**
 * Parse an RRULE line into a structured object (RFC 5545 §3.3.10).
 *
 * Input: `RRULE:FREQ=WEEKLY;BYDAY=TH;COUNT=10`
 * Output: { freq:'WEEKLY', byDay:['TH'], count:10, raw:'FREQ=WEEKLY;BYDAY=TH;COUNT=10' }
 */
export function parseRecurrence(vevent: string): ParsedRecurrence | undefined {
  const raw = parseICalValue(vevent, 'RRULE');
  if (!raw) return undefined;
  const parts = raw.split(';');
  const result: ParsedRecurrence = { raw };
  for (const part of parts) {
    const [k, v] = part.split('=');
    if (!k || !v) continue;
    switch (k.toUpperCase()) {
      case 'FREQ': result.freq = v.toUpperCase(); break;
      case 'INTERVAL': result.interval = parseInt(v, 10); break;
      case 'COUNT': result.count = parseInt(v, 10); break;
      case 'UNTIL': result.until = v; break;
      case 'BYDAY': result.byDay = v.split(','); break;
      case 'BYMONTHDAY': result.byMonthDay = v.split(',').map(n => parseInt(n, 10)); break;
    }
  }
  return result;
}

/**
 * Parse the ORGANIZER line (RFC 5545 §3.8.4.3).
 * Returns { email, name? } or undefined when no ORGANIZER is present.
 */
export function parseOrganizer(vevent: string): { email: string; name?: string } | undefined {
  const orgLine = vevent.match(/^ORGANIZER[^:]*:mailto:(.+)$/im);
  if (!orgLine) return undefined;
  const email = orgLine[1].trim();
  const cnMatch = orgLine[0].match(/CN="?([^";:]+)"?/i);
  return { email, name: cnMatch ? cnMatch[1].trim() : undefined };
}

/**
 * Parse CATEGORIES (RFC 5545 §3.8.1.2). Values are comma-separated.
 */
export function parseCategories(vevent: string): string[] | undefined {
  const raw = parseICalValue(vevent, 'CATEGORIES');
  if (!raw) return undefined;
  return raw.split(',').map(c => c.trim()).filter(Boolean);
}

/**
 * Build a VALARM ICS block (RFC 5545 §3.6.6).
 *
 * @param trigger Duration string (e.g. '-PT15M', '-P1D', '-PT1H') or absolute datetime.
 * @param action  'DISPLAY' (default) or 'EMAIL'.
 * @param description Optional description shown in the alarm popup.
 * @returns A complete BEGIN:VALARM…END:VALARM block (without trailing CRLF).
 */
export function buildValarm(
  trigger: string,
  action: string = 'DISPLAY',
  description?: string,
): string {
  const lines = [
    'BEGIN:VALARM',
    `ACTION:${action.toUpperCase()}`,
    `TRIGGER:${trigger}`,
  ];
  if (description) {
    lines.push(`DESCRIPTION:${escapeICalText(description)}`);
  } else if (action.toUpperCase() === 'DISPLAY') {
    lines.push('DESCRIPTION:Reminder');
  }
  lines.push('END:VALARM');
  return lines.join('\r\n');
}

/**
 * Build an RRULE value string from structured params (RFC 5545 §3.3.10).
 *
 * @param recurrence { freq, interval?, until?, count?, byDay?, byMonthDay? }
 * @returns e.g. `FREQ=WEEKLY;BYDAY=TH;COUNT=10`
 */
export function buildRrule(recurrence: {
  freq: string;
  interval?: number;
  until?: string;
  count?: number;
  byDay?: string[];
  byMonthDay?: number[];
}): string {
  const parts: string[] = [`FREQ=${recurrence.freq.toUpperCase()}`];
  if (recurrence.interval) parts.push(`INTERVAL=${recurrence.interval}`);
  if (recurrence.count !== undefined) parts.push(`COUNT=${recurrence.count}`);
  if (recurrence.until) parts.push(`UNTIL=${recurrence.until}`);
  if (recurrence.byDay?.length) parts.push(`BYDAY=${recurrence.byDay.join(',')}`);
  if (recurrence.byMonthDay?.length) parts.push(`BYMONTHDAY=${recurrence.byMonthDay.join(',')}`);
  return parts.join(';');
}

/**
 * Extract the VEVENT block from iCalendar data.
 * This avoids matching properties from VTIMEZONE or other components.
 */
export function extractVEvent(data: string): string {
  const match = data.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/);
  return match ? match[0] : data;
}

/**
 * Parse an iCalendar property value from within a VEVENT block.
 * Handles simple (KEY:value), parameterized (KEY;TZID=...:value),
 * and VALUE=DATE (KEY;VALUE=DATE:20260319) forms.
 * Also handles line folding (continuation lines starting with space/tab).
 */
export function parseICalValue(vevent: string, key: string): string | undefined {
  // Match KEY followed by either ; (params) or : (value), capturing the rest
  const regex = new RegExp(`^(${key}[;:].*)$`, 'm');
  const match = vevent.match(regex);
  if (!match) return undefined;

  // Handle line folding: continuation lines start with space or tab
  let fullLine = match[1];
  const lines = vevent.split(/\r?\n/);
  const matchIdx = lines.findIndex(l => l === fullLine || l.startsWith(fullLine));
  if (matchIdx >= 0) {
    for (let i = matchIdx + 1; i < lines.length; i++) {
      if (lines[i].startsWith(' ') || lines[i].startsWith('\t')) {
        fullLine += lines[i].substring(1);
      } else {
        break;
      }
    }
  }

  // Extract the value after the last colon in the property line
  // For DTSTART;TZID=Europe/Rome:20260320T083000 → 20260320T083000
  // For DTSTART:20220210T154500Z → 20220210T154500Z
  // For DTSTART;VALUE=DATE:20260324 → 20260324
  const colonIdx = fullLine.indexOf(':');
  if (colonIdx === -1) return undefined;
  return fullLine.substring(colonIdx + 1).trim();
}

/**
 * Format an iCalendar date/datetime string to ISO 8601.
 * Input formats: 20260320T083000, 20260320T083000Z, 20260324
 * Output: 2026-03-20T08:30:00, 2026-03-20T08:30:00Z, 2026-03-24
 */
export function formatICalDate(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const cleaned = raw.replace(/\r/g, '');

  // All-day date: 20260324 (8 digits)
  if (/^\d{8}$/.test(cleaned)) {
    return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}`;
  }

  // DateTime: 20260320T083000 or 20260320T083000Z
  const dtMatch = cleaned.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
  if (dtMatch) {
    const [, y, m, d, hh, mm, ss, z] = dtMatch;
    return `${y}-${m}-${d}T${hh}:${mm}:${ss}${z}`;
  }

  return cleaned;
}

export function parseCalendarObject(obj: DAVCalendarObject): CalendarEvent {
  const vevent = extractVEvent(obj.data || '');
  const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
  const description = parseICalValue(vevent, 'DESCRIPTION');
  const rawStart = parseICalValue(vevent, 'DTSTART');
  const rawEnd = parseICalValue(vevent, 'DTEND');
  const location = parseICalValue(vevent, 'LOCATION');
  const uid = parseICalValue(vevent, 'UID') || obj.url || '';

  const base: CalendarEvent = {
    id: uid,
    url: obj.url || '',
    title: unescapeICalText(title),
    description: description ? unescapeICalText(description) : undefined,
    start: formatICalDate(rawStart),
    end: formatICalDate(rawEnd),
    location: location ? unescapeICalText(location) : undefined,
    attendees: parseAttendees(vevent),
    ...parseExtendedProperties(vevent),
  };
  // Why: Localize at the point of construction so EVERY caller of
  // parseCalendarObject (getCalendarEventById, the unbounded fallback in
  // getCalendarEvents, the post-update return in updateCalendarEvent) gets
  // the same TZ-aware shape. The agent's morning/afternoon heuristic only
  // works correctly when the JSON it sees is in the user's local zone.
  return localizeEventTimes(base);
}

/**
 * Extract all RFC 5545 properties beyond the baseline set (SUMMARY, DTSTART,
 * etc.) from a VEVENT block. Centralised so both parseCalendarObject and
 * parseExpandedMultistatus surface the same fields.
 *
 * Why: Previously, alarms, recurrence, status, transparency, priority,
 * categories, classification, organizer, and sequence were silently
 * discarded. This makes them visible to the agent on every read path.
 */
function parseExtendedProperties(vevent: string): Partial<CalendarEvent> {
  const isAllDay = /^DTSTART[;:]VALUE=DATE/m.test(vevent);
  const status = parseICalValue(vevent, 'STATUS');
  const transparency = parseICalValue(vevent, 'TRANSP');
  const priorityRaw = parseICalValue(vevent, 'PRIORITY');
  const classification = parseICalValue(vevent, 'CLASS');
  const sequenceRaw = parseICalValue(vevent, 'SEQUENCE');
  const eventUrl = parseICalValue(vevent, 'URL');

  return {
    allDay: isAllDay || undefined,
    status: status || undefined,
    transparency: transparency || undefined,
    priority: priorityRaw !== undefined ? parseInt(priorityRaw, 10) : undefined,
    classification: classification || undefined,
    categories: parseCategories(vevent),
    alarms: parseAlarms(vevent),
    recurrence: parseRecurrence(vevent),
    organizer: parseOrganizer(vevent),
    sequence: sequenceRaw !== undefined ? parseInt(sequenceRaw, 10) : undefined,
    ...(eventUrl ? { url: eventUrl } : {}),
  };
}

/**
 * Decode the small set of XML entities that may appear inside a
 * <C:calendar-data> CDATA block. Most servers wrap calendar-data in CDATA
 * (so no decoding needed), but some encode `<`, `>`, `&` inline.
 */
export function decodeXMLEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/**
 * Parse a CalDAV multistatus REPORT response that contains expanded VEVENTs.
 *
 * Why: After issuing a calendar-query with `<C:expand>`, the server returns one
 * <D:response> per source calendar object, each containing a <C:calendar-data>
 * CDATA block with one VEVENT per expanded occurrence. We need to extract every
 * VEVENT — not just the first — because a single recurring event yields N
 * occurrences in the window.
 *
 * Note: some events span multiple <D:response> blocks (one per original ics
 * file), and within each block the VCALENDAR may contain multiple VEVENTs (one
 * per occurrence within that recurrence series).
 */
export function parseExpandedMultistatus(
  xml: string,
  calendarUrl: string,
): CalendarEvent[] {
  const events: CalendarEvent[] = [];

  // Match every <D:response>...</D:response> block. We then extract the href and
  // the calendar-data within, then split out each VEVENT inside the iCal data.
  // Be permissive with namespace prefixes (D:, C:, default ns).
  const responseBlockRe = /<(?:[A-Za-z0-9]+:)?response\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?response>/gi;
  const hrefRe = /<(?:[A-Za-z0-9]+:)?href\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?href>/i;
  // calendar-data may use CDATA or be entity-encoded
  const calDataRe = /<(?:[A-Za-z0-9]+:)?calendar-data\b[^>]*>(?:<!\[CDATA\[([\s\S]*?)\]\]>|([\s\S]*?))<\/(?:[A-Za-z0-9]+:)?calendar-data>/i;

  let blockMatch: RegExpExecArray | null;
  while ((blockMatch = responseBlockRe.exec(xml)) !== null) {
    const block = blockMatch[1];
    const hrefMatch = block.match(hrefRe);
    const href = hrefMatch ? hrefMatch[1].trim() : '';
    // Construct an absolute URL for the source ics file. The href is normally
    // a server-relative path (e.g. /dav/calendars/.../event.ics). Resolve it
    // against the calendar URL when that's absolute; otherwise return as-is.
    let objectUrl = '';
    if (href) {
      try {
        objectUrl = new URL(href, calendarUrl).toString();
      } catch {
        objectUrl = href;
      }
    }

    const dataMatch = block.match(calDataRe);
    if (!dataMatch) continue;
    const ical = (dataMatch[1] ?? dataMatch[2] ?? '').replace(/^\s+|\s+$/g, '');
    if (!ical) continue;
    const decoded = decodeXMLEntities(ical);

    // Each VCALENDAR may contain multiple VEVENTs (one per expanded occurrence).
    const veventRe = /BEGIN:VEVENT[\s\S]*?END:VEVENT/g;
    let vMatch: RegExpExecArray | null;
    while ((vMatch = veventRe.exec(decoded)) !== null) {
      const vevent = vMatch[0];
      const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
      const description = parseICalValue(vevent, 'DESCRIPTION');
      const rawStart = parseICalValue(vevent, 'DTSTART');
      const rawEnd = parseICalValue(vevent, 'DTEND');
      const location = parseICalValue(vevent, 'LOCATION');
      const uid = parseICalValue(vevent, 'UID') || objectUrl || '';
      const recurrenceId = parseICalValue(vevent, 'RECURRENCE-ID');

      // Disambiguate occurrences of the same UID by appending the RECURRENCE-ID
      // so callers (or downstream get_calendar_event) can distinguish them.
      const id = recurrenceId ? `${uid}_${recurrenceId}` : uid;

      // Why: Localize each occurrence so the agent sees Bunge's 13:00Z
      // start as `09:00:00-04:00` (EDT) and correctly classifies it as
      // morning. The original UTC instant is preserved in startUtc/endUtc.
      events.push(
        localizeEventTimes({
          id,
          url: objectUrl,
          title: unescapeICalText(title),
          description: description ? unescapeICalText(description) : undefined,
          start: formatICalDate(rawStart),
          end: formatICalDate(rawEnd),
          location: location ? unescapeICalText(location) : undefined,
          attendees: parseAttendees(vevent),
          ...parseExtendedProperties(vevent),
        }),
      );
    }
  }

  return events;
}

/** UUID pattern: 8-4-4-4-12 hex digits. */
const UUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

/**
 * Normalize a calendarId input to the canonical CalDAV collection URL form
 * (full URL with trailing slash) accepted by the Fastmail CalDAV server.
 *
 * Accepts three forms:
 *  1. Bare UUID:                  `4c646201-472c-...`
 *     → `https://caldav.fastmail.com/dav/calendars/user/<username>/<uuid>/`
 *  2. Full URL without trailing slash: `https://caldav.fastmail.com/.../uuid`
 *     → appends `/`
 *  3. Full URL with trailing slash: `https://caldav.fastmail.com/.../uuid/`
 *     → returned unchanged
 *
 * Any other input (non-empty, non-UUID, non-URL) throws a descriptive error.
 * Empty string throws immediately.
 *
 * Why: list_calendar_events previously matched by display name or exact URL,
 * which happened to work when the model passed the full URL. create_calendar_event
 * had trailing-slash normalization but no UUID expansion. Extracting this shared
 * helper makes all calendarId-accepting methods symmetric.
 */
export function normalizeCalendarId(input: string, username: string): string {
  if (!input || input.trim().length === 0) {
    throw new Error('calendarId must not be empty');
  }
  const trimmed = input.trim();

  // Form 1: bare UUID → expand to full CalDAV URL
  if (UUID_RE.test(trimmed)) {
    return `https://caldav.fastmail.com/dav/calendars/user/${username}/${trimmed}/`;
  }

  // Forms 2 & 3: full URL — just ensure trailing slash
  if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
    return trimmed.endsWith('/') ? trimmed : trimmed + '/';
  }

  throw new Error(
    `calendarId must be a UUID (e.g. "4c646201-472c-...") or a full CalDAV URL ` +
    `(e.g. "https://caldav.fastmail.com/dav/calendars/user/..."), got: "${trimmed}"`
  );
}

/**
 * Unescape an iCalendar text value (RFC 5545 §3.3.11).
 * Reverses escaping of newlines, semicolons, commas, and backslashes.
 */
export function unescapeICalText(value: string): string {
  return value
    .replace(/\\n/gi, '\n')
    .replace(/\\;/g, ';')
    .replace(/\\,/g, ',')
    .replace(/\\\\/g, '\\');
}

/**
 * Escape a text value for use in an iCalendar property (RFC 5545 §3.3.11).
 * Backslashes, newlines, commas, and semicolons must be escaped.
 */
export function escapeICalText(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\r?\n/g, '\\n');
}

/**
 * Fetch events from multiple calendars in parallel, merge, sort, and slice.
 *
 * Why: sequential per-calendar iteration with an early-exit
 * `if (allEvents.length >= limit) break` causes calendar starvation — a
 * volume-heavy calendar iterated first (e.g. a daily-supplement "TRT" calendar)
 * fills the limit budget, so calendars iterated later (where the user's
 * Thursday "Bunge" meeting lives) are never queried at all. The fix is to
 * (a) query every calendar with URL set in parallel via Promise.allSettled,
 * (b) cap each calendar's contribution to limit*2 (safety bound against a
 *     runaway calendar OOMing the merge),
 * (c) merge, sort by start ASC, then slice to the global limit so the
 *     low-volume calendar's events still get a fair shot at the result set.
 *
 * Promise.allSettled (vs Promise.all) means a single bad calendar (network
 * error, 4xx/5xx) doesn't tank the whole request — failures get logged with
 * the supplied prefix and the calendar's display name (or URL fallback), and
 * the caller still receives events from every other calendar that succeeded.
 *
 * Both the bounded path (server-side <C:expand> REPORT) and the unbounded
 * fallback (tsdav fetchCalendarObjects) call into this so the starvation fix
 * applies uniformly. Commit 9f3160c originally fixed only the bounded path;
 * the critic's anchoring-check found the unbounded path was still vulnerable.
 *
 * @param targetCalendars  Calendars to query. Calendars without a URL are skipped.
 * @param fetchFn          Per-calendar fetch implementation. Returns events.
 * @param limit            Global slice ceiling on the merged result.
 * @param logPrefix        Prefix used in console.error for rejected per-calendar promises.
 * @returns                Merged, sorted (start ASC), and globally sliced events.
 */
export async function parallelFetchAndMerge(
  targetCalendars: DAVCalendar[],
  fetchFn: (cal: DAVCalendar) => Promise<CalendarEvent[]>,
  limit: number,
  logPrefix: string,
): Promise<CalendarEvent[]> {
  const calendarsWithUrl = targetCalendars.filter(cal => !!cal.url);
  const perCalendarCeiling = Math.max(limit * 2, 1);

  const calendarResults = await Promise.allSettled(
    calendarsWithUrl.map(async cal => {
      const events = await fetchFn(cal);
      return { cal, events: events.slice(0, perCalendarCeiling) };
    }),
  );

  const allEvents: CalendarEvent[] = [];
  const warnings: string[] = [];
  for (let i = 0; i < calendarResults.length; i++) {
    const result = calendarResults[i];
    if (result.status === 'fulfilled') {
      allEvents.push(...result.value.events);
    } else {
      const cal = calendarsWithUrl[i];
      const label = cal?.displayName ? String(cal.displayName) : (cal?.url || `calendar #${i}`);
      const reason = result.reason instanceof Error ? result.reason.message : String(result.reason);
      console.error(`${logPrefix} for calendar "${label}": ${reason}`);
      warnings.push(label);
    }
  }
  allEvents.sort((a, b) => (a.start || '').localeCompare(b.start || ''));
  return allEvents.slice(0, limit);
}

export class CalDAVCalendarClient {
  private config: CalDAVConfig;
  private client: DAVClient | null = null;
  private calendars: DAVCalendar[] | null = null;

  constructor(config: CalDAVConfig) {
    this.config = config;
  }

  private async getClient(): Promise<DAVClient> {
    if (this.client) return this.client;

    this.client = new DAVClient({
      serverUrl: this.config.serverUrl || 'https://caldav.fastmail.com',
      credentials: {
        username: this.config.username,
        password: this.config.password,
      },
      authMethod: 'Basic',
      defaultAccountType: 'caldav',
    });

    await this.client.login();
    return this.client;
  }

  async getCalendars(): Promise<CalendarInfo[]> {
    const client = await this.getClient();
    const calendars = await client.fetchCalendars();
    this.calendars = calendars;

    return calendars
      .filter(c => c.displayName !== 'DEFAULT_TASK_CALENDAR_NAME')
      .map(c => ({
        id: c.url || '',
        displayName: String(c.displayName || 'Unnamed'),
        url: c.url || '',
        description: c.description || undefined,
        color: (c as any).calendarColor || undefined,
      }));
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50, startDate?: string, endDate?: string): Promise<CalendarEvent[]> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    let targetCalendars = this.calendars.filter(
      c => c.displayName !== 'DEFAULT_TASK_CALENDAR_NAME'
    );
    if (calendarId) {
      // Normalize so bare UUIDs and slash-variants all resolve to the same
      // canonical URL form, making calendarId filtering symmetric with create.
      const normalizedId = normalizeCalendarId(calendarId, this.config.username);
      targetCalendars = targetCalendars.filter(
        c => c.url === normalizedId || c.displayName === calendarId
      );
    }

    // When ANY time bound is provided, use server-side recurrence expansion
    // via CalDAV REPORT with <C:expand> (RFC 4791 §9.6.5). Without this, the
    // server returns the recurrence master VEVENT with its original DTSTART,
    // making recurring events invisible to time-windowed queries.
    //
    // For single-bound queries we synthesize the missing bound with sane defaults:
    //   - Only startDate: end = start + 90 days (covers ~3 months forward)
    //   - Only endDate:   start = end - 30 days (covers ~1 month backward)
    //
    // TZ resolution: every raw input flows through resolveDateInput() so that
    // naive datetimes (no Z, no offset) are interpreted as wall-clock time in
    // the user's configured timezone. Without this, the agent's "Thursday
    // morning" (passed as 2026-05-07T00:00:00 → 2026-05-07T12:00:00) gets
    // interpreted as a UTC window, missing the user's 09:00 EDT meeting at
    // 13:00Z. After resolution everything downstream is UTC ISO 8601, so the
    // existing 90/30-day arithmetic and CalDAV formatting are unaffected.
    if (startDate || endDate) {
      const MS_PER_DAY = 86400000;
      let effectiveStart: string;
      let effectiveEnd: string;

      // Normalize provided inputs first — surface any parse failures with
      // the same "Invalid startDate"/"Invalid endDate" prefixes the existing
      // tests rely on.
      let normalizedStart: string | undefined;
      let normalizedEnd: string | undefined;
      if (startDate) {
        try {
          normalizedStart = resolveDateInput(startDate);
        } catch (err: any) {
          throw new Error(`Invalid startDate: ${startDate} (${err?.message || err})`);
        }
      }
      if (endDate) {
        try {
          normalizedEnd = resolveDateInput(endDate);
        } catch (err: any) {
          throw new Error(`Invalid endDate: ${endDate} (${err?.message || err})`);
        }
      }

      if (normalizedStart && normalizedEnd) {
        effectiveStart = normalizedStart;
        effectiveEnd = normalizedEnd;
      } else if (normalizedStart) {
        effectiveStart = normalizedStart;
        const startDt = new Date(normalizedStart);
        if (isNaN(startDt.getTime())) {
          throw new Error(`Invalid startDate: ${startDate}`);
        }
        effectiveEnd = new Date(startDt.getTime() + 90 * MS_PER_DAY).toISOString();
      } else {
        effectiveEnd = normalizedEnd!;
        const endDt = new Date(normalizedEnd!);
        if (isNaN(endDt.getTime())) {
          throw new Error(`Invalid endDate: ${endDate}`);
        }
        effectiveStart = new Date(endDt.getTime() - 30 * MS_PER_DAY).toISOString();
      }

      // Query all calendars in parallel. Why: sequential iteration with an
      // early-exit `if (allEvents.length >= limit) break` causes calendar
      // starvation — a volume-heavy calendar (e.g. one daily-supplement
      // "TRT" calendar emitting ~66 events/window) can fill the limit budget
      // on the first iteration, so calendars iterated later (where the user's
      // recurring "Bunge" Thursday meeting actually lives) are never queried.
      // Use Promise.allSettled so a single bad calendar (network error, 4xx/5xx)
      // doesn't fail the whole request — log and continue, surfacing failed
      // calendar names in a warnings array on the calendar-bound path.
      // Per-calendar fetch ceiling: limit*2 is a safety bound to prevent one
      // runaway calendar from OOMing us, while still leaving headroom for the
      // global merge+slice to surface events from less-prolific calendars.
      return parallelFetchAndMerge(
        targetCalendars,
        cal => this.fetchExpandedEvents(cal.url!, effectiveStart, effectiveEnd),
        limit,
        '[caldav] expand REPORT failed',
      );
    }

    // Fallback: neither bound provided — unbounded query.
    // Recurring events will only return their master DTSTART.
    // Why: even on the unbounded path, sequential iteration with the previous
    // `if (allEvents.length >= limit) break` early-exit caused calendar
    // starvation — a volume-heavy calendar iterated first (e.g. TRT supplements)
    // would saturate the limit budget and prevent later calendars from being
    // queried, so a recurring "Bunge" meeting in a less-prolific calendar would
    // never surface. Use the same Promise.allSettled + merge + slice pattern as
    // the bounded path. (Commit 9f3160c only fixed the bounded branch; this
    // closes the latent unbounded-path starvation the critic flagged.)
    const fetchOptions: any = {};
    return parallelFetchAndMerge(
      targetCalendars,
      async cal => {
        const objects = await client.fetchCalendarObjects({ calendar: cal, ...fetchOptions });
        return objects.map(obj => parseCalendarObject(obj));
      },
      limit,
      '[caldav] fetchCalendarObjects failed',
    );
  }

  /**
   * Fetch events from a calendar with server-side recurrence expansion.
   *
   * Why: tsdav's `fetchCalendarObjects` does not include `<C:expand>` in the
   * REPORT body, so recurring events are returned as the recurrence master with
   * its original DTSTART. This breaks time-windowed queries — a weekly event
   * created in 2024 won't show up in a 2026 window.
   *
   * What: Issues a raw CalDAV REPORT (`calendar-query`) with both a
   * `<C:time-range>` filter AND `<C:expand>` in the requested calendar-data
   * property. The server returns one VEVENT per occurrence in the window, each
   * with its expanded DTSTART and a RECURRENCE-ID identifying the instance.
   *
   * RFC 4791 §9.6.5: <C:expand> "MUST be transformed by the server into a
   * collection of one or more VEVENT components, one per recurrence instance".
   */
  private async fetchExpandedEvents(
    calendarUrl: string,
    startISO: string,
    endISO: string,
  ): Promise<CalendarEvent[]> {
    const startCal = toCalDateUTC(startISO);
    const endCal = toCalDateUTC(endISO);

    const reportBody = `<?xml version="1.0" encoding="UTF-8"?>
<C:calendar-query xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:prop>
    <D:getetag/>
    <C:calendar-data>
      <C:expand start="${startCal}" end="${endCal}"/>
    </C:calendar-data>
  </D:prop>
  <C:filter>
    <C:comp-filter name="VCALENDAR">
      <C:comp-filter name="VEVENT">
        <C:time-range start="${startCal}" end="${endCal}"/>
      </C:comp-filter>
    </C:comp-filter>
  </C:filter>
</C:calendar-query>`;

    const auth = Buffer.from(
      `${this.config.username}:${this.config.password}`,
    ).toString('base64');
    // Ensure trailing slash on collection URL (CalDAV convention)
    const url = calendarUrl.endsWith('/') ? calendarUrl : calendarUrl + '/';

    const response = await fetch(url, {
      method: 'REPORT',
      headers: {
        Authorization: `Basic ${auth}`,
        Depth: '1',
        'Content-Type': 'application/xml; charset=utf-8',
      },
      body: reportBody,
    });

    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(
        `CalDAV expand REPORT failed: ${response.status} ${response.statusText}: ${body.slice(0, 300)}`,
      );
    }

    const xml = await response.text();
    return parseExpandedMultistatus(xml, calendarUrl);
  }

  async getCalendarEventById(eventId: string): Promise<CalendarEvent | null> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    for (const cal of this.calendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal });
      for (const obj of objects) {
        const vevent = extractVEvent(obj.data || '');
        const uid = parseICalValue(vevent, 'UID');
        if (uid === eventId || obj.url === eventId) {
          return parseCalendarObject(obj);
        }
      }
    }

    return null;
  }

  /**
   * Update an existing CalDAV calendar event by UID.
   * Why: Allows patching a subset of event fields without replacing the entire event.
   * What: GETs the raw ICS for the event, replaces only the provided fields in the
   *       VEVENT block, then PUTs the modified ICS back to the same URL.
   * Test: Create an event, call updateCalendarEvent with a new title, then
   *       getCalendarEventById and assert the title changed while other fields are preserved.
   */
  async updateCalendarEvent(eventId: string, updates: {
    title?: string;
    description?: string;
    start?: string;
    end?: string;
    location?: string;
    participants?: string[];
    attachments?: string[];
    /** Alarms to set on the event (replaces existing alarms). */
    alarms?: Array<{ trigger: string; action?: string; description?: string }>;
    /** Event status: TENTATIVE, CONFIRMED, or CANCELLED. */
    status?: string;
    /** Transparency: OPAQUE (busy) or TRANSPARENT (free). */
    transparency?: string;
    /** Priority 0-9 (1=highest). */
    priority?: number;
    /** Privacy class: PUBLIC, PRIVATE, or CONFIDENTIAL. */
    classification?: string;
    /** Categories/tags (replaces existing). */
    categories?: string[];
  }): Promise<CalendarEvent> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    // Find the calendar object matching the eventId (UID or URL)
    let targetObj: DAVCalendarObject | null = null;
    for (const cal of this.calendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal });
      for (const obj of objects) {
        const vevent = extractVEvent(obj.data || '');
        const uid = parseICalValue(vevent, 'UID');
        if (uid === eventId || obj.url === eventId) {
          targetObj = obj;
          break;
        }
      }
      if (targetObj) break;
    }

    if (!targetObj) {
      throw new Error(`Calendar event not found: ${eventId}`);
    }

    const originalData = targetObj.data || '';

    // Replace or set a property within the VEVENT block.
    // If the property exists, replace it; otherwise insert before END:VEVENT.
    const setVEventProp = (ics: string, key: string, value: string): string => {
      // Remove any existing line(s) for this key (including folded continuations)
      const keyPattern = new RegExp(`^${key}[;:].*(?:\\r?\\n[ \\t].*)*`, 'gm');
      if (keyPattern.test(ics)) {
        return ics.replace(keyPattern, `${key}:${value}`);
      }
      // Insert before END:VEVENT
      return ics.replace(/END:VEVENT/, `${key}:${value}\r\nEND:VEVENT`);
    };

    let updatedIcs = originalData;

    if (updates.title !== undefined && updates.title !== null && updates.title !== '') {
      updatedIcs = setVEventProp(updatedIcs, 'SUMMARY', escapeICalText(updates.title));
    }
    if (updates.description !== undefined && updates.description !== null && updates.description !== '') {
      updatedIcs = setVEventProp(updatedIcs, 'DESCRIPTION', escapeICalText(updates.description));
    }
    if (updates.location !== undefined && updates.location !== null && updates.location !== '') {
      updatedIcs = setVEventProp(updatedIcs, 'LOCATION', escapeICalText(updates.location));
    }
    if (updates.start !== undefined && updates.start !== null && updates.start !== '') {
      // Why: Replace any TZID-parameterized form with a UTC value to remove
      // ambiguity, and run the input through resolveDateInput so naive datetimes
      // are interpreted in the user's configured zone (not UTC) before being
      // formatted as RFC 5545 UTC. The previous string-mangle stripped offsets
      // entirely, breaking timezone-explicit inputs.
      const dtstartVal = toCalDateUTC(resolveDateInput(updates.start));
      updatedIcs = updatedIcs.replace(
        /^DTSTART[^\r\n]*(\r?\n[ \t][^\r\n]*)*/m,
        `DTSTART:${dtstartVal}`
      );
    }
    if (updates.end !== undefined && updates.end !== null && updates.end !== '') {
      const dtendVal = toCalDateUTC(resolveDateInput(updates.end));
      updatedIcs = updatedIcs.replace(
        /^DTEND[^\r\n]*(\r?\n[ \t][^\r\n]*)*/m,
        `DTEND:${dtendVal}`
      );
    }
    if (updates.participants !== undefined) {
      // Remove all existing ATTENDEE lines
      updatedIcs = updatedIcs.replace(/^ATTENDEE[^\r\n]*(\r?\n[ \t][^\r\n]*)*/gm, '');
      // Remove blank lines left behind
      updatedIcs = updatedIcs.replace(/(\r?\n){2,}/g, '\r\n');
      // Insert new ATTENDEE lines before END:VEVENT. RSVP=TRUE is required so
      // Fastmail's CalDAV server generates and sends iTIP (RFC 5546) invites.
      const attendeeLines = updates.participants
        .map(email => `ATTENDEE;CN="${email}";RSVP=TRUE:mailto:${email}`)
        .join('\r\n');
      if (attendeeLines) {
        updatedIcs = updatedIcs.replace(/END:VEVENT/, `${attendeeLines}\r\nEND:VEVENT`);
        // A scheduling event needs METHOD:REQUEST (VCALENDAR scope) and an
        // ORGANIZER (VEVENT scope) for invites to go out. When updating an
        // event that previously had no attendees these lines are absent, so
        // inject them. Idempotent — skipped when already present.
        if (!/^METHOD:REQUEST\r?\n/m.test(updatedIcs)) {
          updatedIcs = updatedIcs.replace(
            /^(BEGIN:VCALENDAR\r?\nVERSION:2\.0\r?\n)/m,
            `$1METHOD:REQUEST\r\n`
          );
        }
        if (!/^ORGANIZER[;:]/m.test(updatedIcs)) {
          updatedIcs = updatedIcs.replace(
            /END:VEVENT/,
            `ORGANIZER;CN=David Gutowsky:mailto:davidgutowsky@fastmail.com\r\nEND:VEVENT`
          );
        }
      } else {
        // Removing all attendees: a METHOD:REQUEST with zero ATTENDEEs is invalid
        // per RFC 5546, so strip METHOD:REQUEST (VCALENDAR) and the ORGANIZER line
        // (VEVENT) as well, leaving a plain non-scheduling event.
        updatedIcs = updatedIcs.replace(/^METHOD:REQUEST\r?\n/m, '');
        updatedIcs = updatedIcs.replace(/^ORGANIZER[^\r\n]*(\r?\n[ \t][^\r\n]*)*\r?\n/m, '');
      }
    }
    if (updates.attachments !== undefined && updates.attachments.length > 0) {
      // Remove existing ATTACH lines
      updatedIcs = updatedIcs.replace(/^ATTACH[^\r\n]*(\r?\n[ \t][^\r\n]*)*/gm, '');
      updatedIcs = updatedIcs.replace(/(\r?\n){2,}/g, '\r\n');
      // Add new ATTACH lines before END:VEVENT
      const fs = await import('fs');
      const attachLines = updates.attachments.map(filePath => {
        const ext = filePath.split('.').pop()?.toLowerCase() ?? 'bin';
        const mimeMap: Record<string, string> = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', pdf: 'application/pdf', gif: 'image/gif' };
        const mime = mimeMap[ext] ?? 'application/octet-stream';
        const data = fs.readFileSync(filePath);
        const b64 = data.toString('base64');
        // Fold base64 at 75 chars per RFC 5545
        const folded = b64.match(/.{1,75}/g)?.join('\r\n ') ?? b64;
        return `ATTACH;ENCODING=BASE64;FMTTYPE=${mime}:\r\n ${folded}`;
      }).join('\r\n');
      updatedIcs = updatedIcs.replace(/END:VEVENT/, `${attachLines}\r\nEND:VEVENT`);
    }

    // --- Alarms (VALARM) ---
    // Replace existing VALARM blocks (if any) with the new set, or add new ones.
    // Why: Previously alarms couldn't be added via update. Now the agent can
    // attach reminders to an existing event (the original use case).
    if (updates.alarms !== undefined) {
      // Strip all existing VALARM blocks from the VEVENT.
      updatedIcs = updatedIcs.replace(/BEGIN:VALARM[\s\S]*?END:VALARM\r?\n?/gi, '');
      // Inject new VALARM blocks before END:VEVENT.
      if (updates.alarms.length > 0) {
        const alarmBlocks = updates.alarms
          .map(a => buildValarm(a.trigger, a.action, a.description))
          .join('\r\n');
        updatedIcs = updatedIcs.replace(/END:VEVENT/, `${alarmBlocks}\r\nEND:VEVENT`);
      }
    }

    // --- Simple scalar properties (STATUS, TRANSP, PRIORITY, CLASS, CATEGORIES) ---
    if (updates.status !== undefined && updates.status !== '') {
      updatedIcs = setVEventProp(updatedIcs, 'STATUS', updates.status.toUpperCase());
    }
    if (updates.transparency !== undefined && updates.transparency !== '') {
      updatedIcs = setVEventProp(updatedIcs, 'TRANSP', updates.transparency.toUpperCase());
    }
    if (updates.priority !== undefined) {
      updatedIcs = setVEventProp(updatedIcs, 'PRIORITY', String(updates.priority));
    }
    if (updates.classification !== undefined && updates.classification !== '') {
      updatedIcs = setVEventProp(updatedIcs, 'CLASS', updates.classification.toUpperCase());
    }
    if (updates.categories !== undefined) {
      // Remove existing CATEGORIES line(s).
      updatedIcs = updatedIcs.replace(/^CATEGORIES[;:].*(?:\r?\n[ \t][^\r\n]*)*/gm, '');
      if (updates.categories.length > 0) {
        const catVal = updates.categories.map(escapeICalText).join(',');
        updatedIcs = updatedIcs.replace(/END:VEVENT/, `CATEGORIES:${catVal}\r\nEND:VEVENT`);
      }
    }

    // --- SEQUENCE bump (RFC 5545 §3.8.7.4) ---
    // Why: Incrementing SEQUENCE on every update signals to calendar clients
    // that the event changed, so they refresh their cached copy. Without this,
    // some clients (Apple Calendar, Outlook) may show stale data.
    const existingSeqMatch = updatedIcs.match(/^SEQUENCE:(\d+)/m);
    const nextSeq = existingSeqMatch ? parseInt(existingSeqMatch[1], 10) + 1 : 0;
    if (existingSeqMatch) {
      updatedIcs = updatedIcs.replace(/^SEQUENCE:\d+/m, `SEQUENCE:${nextSeq}`);
    } else {
      updatedIcs = updatedIcs.replace(/END:VEVENT/, `SEQUENCE:${nextSeq}\r\nEND:VEVENT`);
    }

    // Bump DTSTAMP to now
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    updatedIcs = updatedIcs.replace(/^DTSTAMP:.*$/m, `DTSTAMP:${now}`);

    await client.updateCalendarObject({
      calendarObject: {
        url: targetObj.url,
        data: updatedIcs,
        etag: targetObj.etag,
      },
    });

    return parseCalendarObject({ ...targetObj, data: updatedIcs });
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string;
    end: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    /** All-day event: emits DTSTART;VALUE=DATE instead of a timed datetime. */
    allDay?: boolean;
    /** Alarms to attach (e.g. [{trigger:'-PT15M'}, {trigger:'-P1D'}]). */
    alarms?: Array<{ trigger: string; action?: string; description?: string }>;
    /** Recurrence rule (e.g. {freq:'WEEKLY', byDay:['TH'], count:10}). */
    recurrence?: {
      freq: string;
      interval?: number;
      until?: string;
      count?: number;
      byDay?: string[];
      byMonthDay?: number[];
    };
    /** Event status: TENTATIVE, CONFIRMED, or CANCELLED. */
    status?: string;
    /** Transparency: OPAQUE (busy) or TRANSPARENT (free). */
    transparency?: string;
    /** Priority 0-9 (1=highest). */
    priority?: number;
    /** Privacy class: PUBLIC, PRIVATE, or CONFIDENTIAL. */
    classification?: string;
    /** Categories/tags. */
    categories?: string[];
    /** Event URL (conference link, etc.). */
    url?: string;
  }): Promise<string> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    // Normalize calendarId: accept bare UUID, full URL with or without trailing
    // slash. normalizeCalendarId expands UUIDs to full CalDAV URLs and ensures
    // the canonical trailing-slash form. Display name is tried as a fallback.
    const normalizedCalId = normalizeCalendarId(event.calendarId, this.config.username);
    const targetCal = this.calendars.find(
      c => c.url === normalizedCalId || c.displayName === event.calendarId
    );
    if (!targetCal) {
      throw new Error(`Calendar not found: ${event.calendarId}`);
    }

    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@fastmail-mcp`;
    const now = toCalDateUTC(new Date().toISOString());
    // Why: A naive event.start like "2026-05-07T09:00:00" should mean 09:00 in
    // the user's local timezone (configured via FASTMAIL_TIMEZONE / TZ), not
    // 09:00 UTC. resolveDateInput() honors explicit Z/offsets and otherwise
    // interprets naive inputs in the configured zone before handing off to
    // toCalDateUTC for RFC 5545 formatting. This complements 9f3160c by
    // making the write path TZ-aware in addition to the read path.
    //
    // For all-day events (allDay=true), we emit VALUE=DATE dates (YYYYMMDD)
    // instead of UTC datetimes, per RFC 5545 §3.3.4.
    const isAllDay = !!event.allDay;
    const dtstart = isAllDay
      ? resolveDateInput(event.start).slice(0, 10).replace(/-/g, '')
      : toCalDateUTC(resolveDateInput(event.start));
    const dtend = isAllDay
      ? resolveDateInput(event.end).slice(0, 10).replace(/-/g, '')
      : toCalDateUTC(resolveDateInput(event.end));
    const dtstartLine = isAllDay ? `DTSTART;VALUE=DATE:${dtstart}` : `DTSTART:${dtstart}`;
    const dtendLine = isAllDay ? `DTEND;VALUE=DATE:${dtend}` : `DTEND:${dtend}`;
    // Why: When attendees are present, the ICS must carry METHOD:REQUEST plus an
    // ORGANIZER and one ATTENDEE line per participant so that Fastmail's CalDAV
    // server generates and sends iTIP (RFC 5546) invite emails. Without these
    // lines the event is created but no invites go out and the attendees never
    // appear on the event. When there are no participants we keep the original
    // minimal ICS (no METHOD/ORGANIZER/ATTENDEE) to avoid spurious scheduling.
    const hasParticipants = !!(event.participants && event.participants.length > 0);
    const attendeeLines = hasParticipants
      ? event.participants!.map(
          p => `ATTENDEE;CN="${p.name || p.email}";RSVP=TRUE:mailto:${p.email}`,
        )
      : [];
    // Build VALARM blocks for each alarm in the request.
    const alarmBlocks = (event.alarms && event.alarms.length > 0)
      ? event.alarms.map(a => buildValarm(a.trigger, a.action, a.description))
      : [];
    // Build RRULE line if recurrence is specified.
    const rruleLine = event.recurrence ? `RRULE:${buildRrule(event.recurrence)}` : '';
    const ical = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      hasParticipants ? 'METHOD:REQUEST' : '',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      dtstartLine,
      dtendLine,
      rruleLine,
      `SUMMARY:${escapeICalText(event.title)}`,
      event.description ? `DESCRIPTION:${escapeICalText(event.description)}` : '',
      event.location ? `LOCATION:${escapeICalText(event.location)}` : '',
      event.status ? `STATUS:${event.status.toUpperCase()}` : '',
      event.transparency ? `TRANSP:${event.transparency.toUpperCase()}` : '',
      event.priority !== undefined ? `PRIORITY:${event.priority}` : '',
      event.classification ? `CLASS:${event.classification.toUpperCase()}` : '',
      event.categories?.length ? `CATEGORIES:${event.categories.map(escapeICalText).join(',')}` : '',
      event.url ? `URL:${event.url}` : '',
      'SEQUENCE:0',
      hasParticipants ? 'ORGANIZER;CN=David Gutowsky:mailto:davidgutowsky@fastmail.com' : '',
      ...attendeeLines,
      ...alarmBlocks,
      'END:VEVENT',
      'END:VCALENDAR',
    ].filter(Boolean).join('\r\n');

    // Issue PUT to CalDAV server. Wrap in try/catch and validate response so
    // that we DO NOT return fake success when the server rejected the event.
    let response: any;
    try {
      response = await client.createCalendarObject({
        calendar: targetCal,
        filename: `${uid}.ics`,
        iCalString: ical,
      });
    } catch (err: any) {
      throw new Error(`CalDAV PUT failed: ${err?.message || String(err)}`);
    }

    // Validate response — tsdav returns a Response (or array of them).
    // Treat any non-ok status as a failure rather than silently returning success.
    const responses = Array.isArray(response) ? response : [response];
    for (const r of responses) {
      if (r && typeof r === 'object' && 'ok' in r && !(r as Response).ok) {
        const status = (r as Response).status;
        let body = '';
        try {
          body = await (r as Response).text();
        } catch {
          // ignore body read failures
        }
        throw new Error(`CalDAV PUT returned ${status}: ${body.slice(0, 500)}`);
      }
    }

    // Verify persistence with a follow-up GET against the calendar object URL.
    // This catches the case where the server returns a 2xx but the event was
    // not actually stored (e.g. silently dropped due to malformed iCal).
    try {
      const verifyUrl = `${targetCal.url.replace(/\/$/, '')}/${uid}.ics`;
      const auth = Buffer.from(
        `${this.config.username}:${this.config.password}`
      ).toString('base64');
      const verifyResp = await fetch(verifyUrl, {
        method: 'GET',
        headers: {
          Authorization: `Basic ${auth}`,
        },
      });
      if (!verifyResp.ok) {
        throw new Error(
          `Event PUT reported success but verification GET returned ${verifyResp.status}. Event was NOT persisted.`
        );
      }
    } catch (err: any) {
      throw new Error(
        `Persistence verification failed: ${err?.message || String(err)}`
      );
    }

    return uid;
  }

  async deleteCalendarEvent(eventId: string, calendarId: string): Promise<void> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    // Normalize calendarId for the same UUID/trailing-slash symmetry as create.
    const targetCals = calendarId
      ? (() => {
          const normalizedId = normalizeCalendarId(calendarId, this.config.username);
          return this.calendars.filter(c => c.url === normalizedId || c.displayName === calendarId);
        })()
      : this.calendars;

    for (const cal of targetCals) {
      const objects = await client.fetchCalendarObjects({ calendar: cal });
      for (const obj of objects) {
        const vevent = extractVEvent(obj.data || '');
        const uid = parseICalValue(vevent, 'UID');
        if (uid === eventId || obj.url === eventId) {
          await client.deleteCalendarObject({ calendarObject: obj });
          return;
        }
      }
    }

    throw new Error(`Event ${eventId} not found in calendar ${calendarId}`);
  }
}
