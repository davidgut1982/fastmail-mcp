import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';
import { resolveDateInput } from './timezone.js';

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

export interface CalendarEvent {
  id: string;
  url: string;
  title: string;
  description?: string;
  start?: string;
  end?: string;
  location?: string;
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

  return {
    id: uid,
    url: obj.url || '',
    title: unescapeICalText(title),
    description: description ? unescapeICalText(description) : undefined,
    start: formatICalDate(rawStart),
    end: formatICalDate(rawEnd),
    location: location ? unescapeICalText(location) : undefined,
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

      events.push({
        id,
        url: objectUrl,
        title: unescapeICalText(title),
        description: description ? unescapeICalText(description) : undefined,
        start: formatICalDate(rawStart),
        end: formatICalDate(rawEnd),
        location: location ? unescapeICalText(location) : undefined,
      });
    }
  }

  return events;
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
      targetCalendars = targetCalendars.filter(
        c => c.url === calendarId || c.displayName === calendarId
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
      const perCalendarCeiling = Math.max(limit * 2, 1);
      const calendarResults = await Promise.allSettled(
        targetCalendars
          .filter(cal => !!cal.url)
          .map(async cal => {
            const expanded = await this.fetchExpandedEvents(cal.url!, effectiveStart, effectiveEnd);
            return { cal, events: expanded.slice(0, perCalendarCeiling) };
          }),
      );

      const allEvents: CalendarEvent[] = [];
      const warnings: string[] = [];
      for (let i = 0; i < calendarResults.length; i++) {
        const result = calendarResults[i];
        if (result.status === 'fulfilled') {
          allEvents.push(...result.value.events);
        } else {
          // Log failed per-calendar query and continue. Identify the calendar
          // by displayName when available, falling back to URL or index.
          const cal = targetCalendars.filter(c => !!c.url)[i];
          const label = cal?.displayName ? String(cal.displayName) : (cal?.url || `calendar #${i}`);
          const reason = result.reason instanceof Error ? result.reason.message : String(result.reason);
          console.error(`[caldav] expand REPORT failed for calendar "${label}": ${reason}`);
          warnings.push(label);
        }
      }
      allEvents.sort((a, b) => (a.start || '').localeCompare(b.start || ''));
      return allEvents.slice(0, limit);
    }

    // Fallback: neither bound provided — unbounded query.
    // Recurring events will only return their master DTSTART.
    const fetchOptions: any = {};

    const allEvents: CalendarEvent[] = [];
    for (const cal of targetCalendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal, ...fetchOptions });
      for (const obj of objects) {
        allEvents.push(parseCalendarObject(obj));
      }
      if (allEvents.length >= limit) break;
    }

    // Sort by start date ascending
    allEvents.sort((a, b) => (a.start || '').localeCompare(b.start || ''));

    return allEvents.slice(0, limit);
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
      // Insert new ATTENDEE lines before END:VEVENT
      const attendeeLines = updates.participants
        .map(email => `ATTENDEE;CN=${email}:mailto:${email}`)
        .join('\r\n');
      if (attendeeLines) {
        updatedIcs = updatedIcs.replace(/END:VEVENT/, `${attendeeLines}\r\nEND:VEVENT`);
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
  }): Promise<string> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    const targetCal = this.calendars.find(
      c => c.url === event.calendarId || c.displayName === event.calendarId
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
    const dtstart = toCalDateUTC(resolveDateInput(event.start));
    const dtend = toCalDateUTC(resolveDateInput(event.end));
    const ical = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `DTSTART:${dtstart}`,
      `DTEND:${dtend}`,
      `SUMMARY:${escapeICalText(event.title)}`,
      event.description ? `DESCRIPTION:${escapeICalText(event.description)}` : '',
      event.location ? `LOCATION:${escapeICalText(event.location)}` : '',
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
}
