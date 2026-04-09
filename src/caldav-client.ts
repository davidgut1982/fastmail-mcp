import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';

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

    const fetchOptions: any = {};
    if (startDate || endDate) {
      fetchOptions.timeRange = {
        start: startDate || '1970-01-01T00:00:00Z',
        end: endDate || '2099-12-31T23:59:59Z',
      };
    }

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

    if (updates.title !== undefined) {
      updatedIcs = setVEventProp(updatedIcs, 'SUMMARY', escapeICalText(updates.title));
    }
    if (updates.description !== undefined) {
      updatedIcs = setVEventProp(updatedIcs, 'DESCRIPTION', escapeICalText(updates.description));
    }
    if (updates.location !== undefined) {
      updatedIcs = setVEventProp(updatedIcs, 'LOCATION', escapeICalText(updates.location));
    }
    if (updates.start !== undefined) {
      // Preserve existing DTSTART params (e.g. TZID) when replacing the value
      const existingDtstart = originalData.match(/^(DTSTART[^:]*):.*$/m);
      const dtstartKey = existingDtstart ? existingDtstart[1] : 'DTSTART';
      const dtstartVal = updates.start.replace(/[-:]/g, '').replace(/\.\d{3}/, '');
      updatedIcs = updatedIcs.replace(
        /^DTSTART[^\r\n]*(\r?\n[ \t][^\r\n]*)*/m,
        `${dtstartKey}:${dtstartVal}`
      );
    }
    if (updates.end !== undefined) {
      const existingDtend = originalData.match(/^(DTEND[^:]*):.*$/m);
      const dtendKey = existingDtend ? existingDtend[1] : 'DTEND';
      const dtendVal = updates.end.replace(/[-:]/g, '').replace(/\.\d{3}/, '');
      updatedIcs = updatedIcs.replace(
        /^DTEND[^\r\n]*(\r?\n[ \t][^\r\n]*)*/m,
        `${dtendKey}:${dtendVal}`
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
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    const ical = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `DTSTART:${event.start.replace(/[-:]/g, '')}`,
      `DTEND:${event.end.replace(/[-:]/g, '')}`,
      `SUMMARY:${escapeICalText(event.title)}`,
      event.description ? `DESCRIPTION:${escapeICalText(event.description)}` : '',
      event.location ? `LOCATION:${escapeICalText(event.location)}` : '',
      'END:VEVENT',
      'END:VCALENDAR',
    ].filter(Boolean).join('\r\n');

    await client.createCalendarObject({
      calendar: targetCal,
      filename: `${uid}.ics`,
      iCalString: ical,
    });

    return uid;
  }
}
