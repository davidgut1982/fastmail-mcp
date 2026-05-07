import { describe, it, mock, before, after } from 'node:test';
import assert from 'node:assert/strict';
import {
  extractVEvent,
  parseICalValue,
  formatICalDate,
  parseCalendarObject,
  parseExpandedMultistatus,
  escapeICalText,
  CalDAVCalendarClient,
} from './caldav-client.js';
import {
  resolveDateInput,
  isForceLocalTimezone,
  __resetTimezoneCacheForTests,
} from './timezone.js';
import { localizeEventTimes } from './event-localizer.js';

describe('extractVEvent', () => {
  it('extracts VEVENT block from iCalendar data', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'DTSTART:19700101T000000',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'SUMMARY:Test Event',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    assert.ok(vevent.includes('SUMMARY:Test Event'));
    assert.ok(vevent.includes('DTSTART;TZID=Europe/Rome:20260320T083000'));
    assert.ok(!vevent.includes('VTIMEZONE'));
    assert.ok(!vevent.includes('TZID:Europe/Rome'));
  });

  it('returns original data when no VEVENT block found', () => {
    const data = 'no vevent here';
    assert.equal(extractVEvent(data), data);
  });

  it('ignores VTIMEZONE DTSTART when extracting VEVENT', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'END:STANDARD',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'SUMMARY:Meeting',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    // Should only have the VEVENT DTSTART, not the VTIMEZONE one
    const dtstartMatches = vevent.match(/DTSTART/g);
    assert.equal(dtstartMatches?.length, 1);
    assert.ok(vevent.includes('20260320T083000'));
  });
});

describe('parseICalValue', () => {
  it('handles simple KEY:value format', () => {
    const vevent = 'SUMMARY:Test Event\nDTSTART:20260320T083000Z';
    assert.equal(parseICalValue(vevent, 'SUMMARY'), 'Test Event');
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000Z');
  });

  it('handles parameterized KEY;TZID=...:value format', () => {
    const vevent = 'DTSTART;TZID=Europe/Rome:20260320T083000\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000');
  });

  it('handles VALUE=DATE format', () => {
    const vevent = 'DTSTART;VALUE=DATE:20260324\nDTEND;VALUE=DATE:20260325';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260324');
    assert.equal(parseICalValue(vevent, 'DTEND'), '20260325');
  });

  it('returns undefined for missing keys', () => {
    const vevent = 'SUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'LOCATION'), undefined);
  });

  it('handles line folding (continuation lines)', () => {
    const vevent = 'DESCRIPTION:This is a long\n description that wraps\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), 'This is a longdescription that wraps');
  });
});

describe('formatICalDate', () => {
  it('formats datetime without timezone', () => {
    assert.equal(formatICalDate('20260320T083000'), '2026-03-20T08:30:00');
  });

  it('formats datetime with Z suffix', () => {
    assert.equal(formatICalDate('20260320T083000Z'), '2026-03-20T08:30:00Z');
  });

  it('formats all-day date', () => {
    assert.equal(formatICalDate('20260324'), '2026-03-24');
  });

  it('returns undefined for undefined input', () => {
    assert.equal(formatICalDate(undefined), undefined);
  });

  it('returns cleaned string for unrecognized formats', () => {
    assert.equal(formatICalDate('something-else'), 'something-else');
  });

  it('strips carriage returns', () => {
    assert.equal(formatICalDate('20260320T083000\r'), '2026-03-20T08:30:00');
  });
});

describe('parseCalendarObject', () => {
  // parseCalendarObject now localizes start/end to the configured timezone
  // (per the localizeEventTimes refactor — see event-localizer.ts). Pin the
  // tests to America/New_York so assertions are deterministic regardless of
  // host system timezone. Save/restore env to avoid bleeding state.
  let savedFmTz: string | undefined;
  let savedTz: string | undefined;
  before(() => {
    savedFmTz = process.env.FASTMAIL_TIMEZONE;
    savedTz = process.env.TZ;
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();
  });
  after(() => {
    if (savedFmTz === undefined) delete process.env.FASTMAIL_TIMEZONE;
    else process.env.FASTMAIL_TIMEZONE = savedFmTz;
    if (savedTz === undefined) delete process.env.TZ;
    else process.env.TZ = savedTz;
    __resetTimezoneCacheForTests();
  });

  it('parses a full calendar object with VTIMEZONE + VEVENT', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'TZOFFSETFROM:+0200',
      'TZOFFSETTO:+0100',
      'END:STANDARD',
      'BEGIN:DAYLIGHT',
      'DTSTART:19700329T020000',
      'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
      'TZOFFSETFROM:+0100',
      'TZOFFSETTO:+0200',
      'END:DAYLIGHT',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:abc123@fastmail',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'DTEND;TZID=Europe/Rome:20260320T093000',
      'SUMMARY:Morning Meeting',
      'DESCRIPTION:Discuss project\\nSecond line',
      'LOCATION:Room A\\, Building 1',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: 'https://caldav.example.com/cal/abc.ics' });

    assert.equal(event.id, 'abc123@fastmail');
    assert.equal(event.url, 'https://caldav.example.com/cal/abc.ics');
    assert.equal(event.title, 'Morning Meeting');
    assert.equal(event.description, 'Discuss project\nSecond line');
    assert.equal(event.location, 'Room A, Building 1');
    // Localizer treats this floating-time value as UTC (DTSTART has no
    // explicit Z but no offset; we don't expand VTIMEZONE), then shifts to
    // America/New_York. 08:30 UTC on 2026-03-20 = 04:30 EDT (DST in effect).
    // The original UTC value is preserved in startUtc/endUtc.
    assert.equal(
      Date.parse(event.start as string),
      Date.parse('2026-03-20T08:30:00Z'),
      `start should represent the same UTC instant, got ${event.start}`,
    );
    assert.equal(event.timezone, 'America/New_York');
    assert.match(event.start as string, /-04:00$/);
    assert.equal(
      Date.parse(event.end as string),
      Date.parse('2026-03-20T09:30:00Z'),
      `end should represent the same UTC instant, got ${event.end}`,
    );
  });

  it('parses an all-day event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday1@fastmail',
      'DTSTART;VALUE=DATE:20260324',
      'DTEND;VALUE=DATE:20260325',
      'SUMMARY:All Day Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-24');
    assert.equal(event.end, '2026-03-25');
    assert.equal(event.title, 'All Day Event');
  });

  it('parses a UTC event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:utc1@fastmail',
      'DTSTART:20260320T083000Z',
      'DTEND:20260320T093000Z',
      'SUMMARY:UTC Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    // After localization: 08:30Z → 04:30 EDT (March is in DST = UTC-4).
    // Original UTC is preserved in startUtc/endUtc.
    assert.equal(
      Date.parse(event.start as string),
      Date.parse('2026-03-20T08:30:00Z'),
      `start should still represent 08:30Z, got ${event.start}`,
    );
    assert.match(event.start as string, /-04:00$/);
    assert.equal(event.startUtc, '2026-03-20T08:30:00Z');
    assert.equal(event.endUtc, '2026-03-20T09:30:00Z');
    assert.equal(event.timezone, 'America/New_York');
  });

  it('defaults title to Untitled when SUMMARY is missing', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:notitle@fastmail',
      'DTSTART:20260320T083000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.title, 'Untitled');
  });

  it('handles missing optional fields', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:minimal@fastmail',
      'DTSTART:20260320T083000Z',
      'SUMMARY:Minimal',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.description, undefined);
    assert.equal(event.location, undefined);
    assert.equal(event.end, undefined);
  });
});

describe('escapeICalText', () => {
  it('escapes backslashes', () => {
    assert.equal(escapeICalText('path\\to\\file'), 'path\\\\to\\\\file');
  });

  it('escapes semicolons', () => {
    assert.equal(escapeICalText('a;b;c'), 'a\\;b\\;c');
  });

  it('escapes commas', () => {
    assert.equal(escapeICalText('Room A, Building 1'), 'Room A\\, Building 1');
  });

  it('escapes newlines', () => {
    assert.equal(escapeICalText('line1\nline2'), 'line1\\nline2');
    assert.equal(escapeICalText('line1\r\nline2'), 'line1\\nline2');
  });

  it('leaves plain text unchanged', () => {
    assert.equal(escapeICalText('Team Standup'), 'Team Standup');
  });

  it('prevents ICS property injection via CRLF', () => {
    const malicious = 'Meeting\r\nATTENDEE:mailto:attacker@evil.com';
    const escaped = escapeICalText(malicious);
    // No literal newlines means the injected ATTENDEE stays inside the text value,
    // not on its own ICS property line
    assert.ok(!escaped.includes('\n'), 'escaped text must not contain literal newlines');
    assert.ok(!escaped.includes('\r'), 'escaped text must not contain literal carriage returns');
    assert.equal(escaped, 'Meeting\\nATTENDEE:mailto:attacker@evil.com');
  });

  it('prevents injection of extra VEVENT components', () => {
    const malicious = 'Meeting\r\nEND:VEVENT\r\nBEGIN:VEVENT\r\nSUMMARY:Injected';
    const escaped = escapeICalText(malicious);
    // No literal newlines means the injected properties stay inside the SUMMARY value
    assert.ok(!escaped.includes('\n'), 'escaped text must not contain literal newlines');
    assert.ok(!escaped.includes('\r'), 'escaped text must not contain literal carriage returns');
    // The entire payload is on one logical ICS line, so END:VEVENT can't terminate the block
    assert.ok(escaped.startsWith('Meeting\\n'), 'newlines should be escaped, not literal');
  });
});

describe('CalDAVCalendarClient.getCalendarEvents', () => {
  // Localizer (event-localizer.ts) rewrites event.start/end into the
  // configured zone with explicit offset. Pin to America/New_York so
  // assertions are deterministic regardless of host TZ.
  let savedFmTz: string | undefined;
  let savedTz: string | undefined;
  before(() => {
    savedFmTz = process.env.FASTMAIL_TIMEZONE;
    savedTz = process.env.TZ;
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();
  });
  after(() => {
    if (savedFmTz === undefined) delete process.env.FASTMAIL_TIMEZONE;
    else process.env.FASTMAIL_TIMEZONE = savedFmTz;
    if (savedTz === undefined) delete process.env.TZ;
    else process.env.TZ = savedTz;
    __resetTimezoneCacheForTests();
  });

  function makeIcal(uid: string, summary: string, dtstart: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTART:${dtstart}`,
      `SUMMARY:${summary}`,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedClient(calendarObjects: Array<{ data: string; url: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    // Override the private getClient method to return a mock DAVClient
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('sorts events by start date ascending', async () => {
    const objects = [
      { data: makeIcal('c@fm', 'Evening', '20260325T200000Z'), url: '/c.ics' },
      { data: makeIcal('a@fm', 'Morning', '20260325T080000Z'), url: '/a.ics' },
      { data: makeIcal('b@fm', 'Afternoon', '20260325T140000Z'), url: '/b.ics' },
    ];
    const { client } = createMockedClient(objects);
    const events = await client.getCalendarEvents(undefined, 50);

    assert.equal(events.length, 3);
    assert.equal(events[0].title, 'Morning');
    assert.equal(events[1].title, 'Afternoon');
    assert.equal(events[2].title, 'Evening');
  });

  it('uses CalDAV REPORT with <C:expand> when both startDate and endDate provided', async () => {
    // When both dates are provided, the client should use server-side recurrence
    // expansion via raw REPORT (not fetchCalendarObjects), so recurring events
    // appear with their actual occurrence DTSTART, not the master DTSTART.
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);

    // Stub global fetch to capture the REPORT request and return a multistatus
    // body with one expanded VEVENT.
    const originalFetch = globalThis.fetch;
    const fetchCalls: any[] = [];
    const expandedXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/a.ics</D:href>
    <D:propstat>
      <D:prop>
        <D:getetag>"abc"</D:getetag>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:a@fm
DTSTART:20260325T100000Z
DTEND:20260325T110000Z
SUMMARY:Event
RECURRENCE-ID:20260325T100000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>
</D:multistatus>`;
    (globalThis as any).fetch = async (url: string, init: any) => {
      fetchCalls.push({ url, init });
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => expandedXml,
      };
    };
    try {
      const events = await client.getCalendarEvents(
        undefined,
        50,
        '2026-03-25T00:00:00Z',
        '2026-03-26T00:00:00Z',
      );

      // fetchCalendarObjects must NOT be called on the time-windowed path
      assert.equal(mockDAVClient.fetchCalendarObjects.mock.calls.length, 0);

      // Exactly one REPORT request issued
      assert.equal(fetchCalls.length, 1);
      assert.equal(fetchCalls[0].init.method, 'REPORT');
      assert.equal(fetchCalls[0].init.headers.Depth, '1');
      // Body contains <C:expand> with the correct CalDAV-formatted dates
      assert.match(fetchCalls[0].init.body, /<C:expand\s+start="20260325T000000Z"\s+end="20260326T000000Z"\s*\/>/);
      assert.match(fetchCalls[0].init.body, /<C:time-range/);

      // Returned event has expanded DTSTART, localized to America/New_York.
      // 10:00Z on 2026-03-25 = 06:00 EDT (DST in effect). Original UTC is
      // preserved in startUtc; timezone field signals locality.
      assert.equal(events.length, 1);
      assert.equal(events[0].title, 'Event');
      assert.equal(
        Date.parse(events[0].start as string),
        Date.parse('2026-03-25T10:00:00Z'),
        `start should represent 10:00Z, got ${events[0].start}`,
      );
      assert.match(events[0].start as string, /-04:00$/);
      assert.equal(events[0].startUtc, '2026-03-25T10:00:00Z');
      assert.equal(events[0].timezone, 'America/New_York');
    } finally {
      (globalThis as any).fetch = originalFetch;
    }
  });

  it('uses <C:expand> when only startDate is provided (synthesizes end = start + 90 days)', async () => {
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260507T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);

    const originalFetch = globalThis.fetch;
    const fetchCalls: any[] = [];
    const expandedXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/a.ics</D:href>
    <D:propstat>
      <D:prop>
        <D:getetag>"abc"</D:getetag>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:a@fm
DTSTART:20260507T100000Z
DTEND:20260507T110000Z
SUMMARY:Event
RECURRENCE-ID:20260507T100000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>
</D:multistatus>`;
    (globalThis as any).fetch = async (url: string, init: any) => {
      fetchCalls.push({ url, init });
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => expandedXml,
      };
    };
    try {
      const events = await client.getCalendarEvents(
        undefined,
        50,
        '2026-03-25T00:00:00Z',
        // No endDate
      );

      // fetchCalendarObjects must NOT be called — should hit expand path
      assert.equal(mockDAVClient.fetchCalendarObjects.mock.calls.length, 0);

      // Exactly one REPORT request issued
      assert.equal(fetchCalls.length, 1);
      assert.equal(fetchCalls[0].init.method, 'REPORT');

      const body = fetchCalls[0].init.body as string;
      // start should match input date
      assert.match(body, /<C:expand\s+start="20260325T000000Z"/);
      // end should be approximately 90 days later (2026-06-23T00:00:00Z)
      const endMatch = body.match(/end="(\d{8}T\d{6}Z)"/);
      assert.ok(endMatch, 'end attribute should be present in expand');
      const endStr = endMatch![1];
      // Parse 20260623T000000Z → epoch ms and check ~90 days from start
      const startMs = Date.parse('2026-03-25T00:00:00Z');
      const endMs = Date.parse(
        `${endStr.slice(0, 4)}-${endStr.slice(4, 6)}-${endStr.slice(6, 8)}T${endStr.slice(9, 11)}:${endStr.slice(11, 13)}:${endStr.slice(13, 15)}Z`,
      );
      const daysDiff = (endMs - startMs) / 86400000;
      assert.ok(
        Math.abs(daysDiff - 90) < 0.01,
        `expected ~90 day window, got ${daysDiff} days`,
      );

      assert.equal(events.length, 1);
      assert.equal(events[0].title, 'Event');
    } finally {
      (globalThis as any).fetch = originalFetch;
    }
  });

  it('uses <C:expand> when only endDate is provided (synthesizes start = end - 30 days)', async () => {
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260507T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);

    const originalFetch = globalThis.fetch;
    const fetchCalls: any[] = [];
    const expandedXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/a.ics</D:href>
    <D:propstat>
      <D:prop>
        <D:getetag>"abc"</D:getetag>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:a@fm
DTSTART:20260507T100000Z
DTEND:20260507T110000Z
SUMMARY:Event
RECURRENCE-ID:20260507T100000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>
</D:multistatus>`;
    (globalThis as any).fetch = async (url: string, init: any) => {
      fetchCalls.push({ url, init });
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => expandedXml,
      };
    };
    try {
      const events = await client.getCalendarEvents(
        undefined,
        50,
        undefined,
        '2026-03-25T00:00:00Z',
      );

      // fetchCalendarObjects must NOT be called — should hit expand path
      assert.equal(mockDAVClient.fetchCalendarObjects.mock.calls.length, 0);

      // Exactly one REPORT request issued
      assert.equal(fetchCalls.length, 1);
      assert.equal(fetchCalls[0].init.method, 'REPORT');

      const body = fetchCalls[0].init.body as string;
      // end should match input date
      assert.match(body, /end="20260325T000000Z"/);
      // start should be approximately 30 days earlier (2026-02-23T00:00:00Z)
      const startMatch = body.match(/<C:expand\s+start="(\d{8}T\d{6}Z)"/);
      assert.ok(startMatch, 'start attribute should be present in expand');
      const startStr = startMatch![1];
      const endMs = Date.parse('2026-03-25T00:00:00Z');
      const startMs = Date.parse(
        `${startStr.slice(0, 4)}-${startStr.slice(4, 6)}-${startStr.slice(6, 8)}T${startStr.slice(9, 11)}:${startStr.slice(11, 13)}:${startStr.slice(13, 15)}Z`,
      );
      const daysDiff = (endMs - startMs) / 86400000;
      assert.ok(
        Math.abs(daysDiff - 30) < 0.01,
        `expected ~30 day window, got ${daysDiff} days`,
      );

      assert.equal(events.length, 1);
      assert.equal(events[0].title, 'Event');
    } finally {
      (globalThis as any).fetch = originalFetch;
    }
  });

  it('falls through to fetchCalendarObjects (legacy path) when neither bound is provided', async () => {
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);

    // Track that fetch (REPORT) is NOT called
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    (globalThis as any).fetch = async () => {
      fetchCalled = true;
      throw new Error('fetch should not be called when no bounds provided');
    };
    try {
      const events = await client.getCalendarEvents(undefined, 50);

      // fetchCalendarObjects MUST be called (legacy path)
      assert.equal(mockDAVClient.fetchCalendarObjects.mock.calls.length, 1);
      // Raw fetch (REPORT/expand) MUST NOT be called
      assert.equal(fetchCalled, false);
      // No timeRange should be passed
      const callArgs = mockDAVClient.fetchCalendarObjects.mock.calls[0].arguments[0];
      assert.equal(callArgs.timeRange, undefined);

      assert.equal(events.length, 1);
      assert.equal(events[0].title, 'Event');
    } finally {
      (globalThis as any).fetch = originalFetch;
    }
  });

  it('throws when only an invalid startDate is provided', async () => {
    const { client } = createMockedClient([]);
    await assert.rejects(
      async () => client.getCalendarEvents(undefined, 50, 'not-a-date'),
      /Invalid startDate/,
    );
  });

  it('throws when only an invalid endDate is provided', async () => {
    const { client } = createMockedClient([]);
    await assert.rejects(
      async () => client.getCalendarEvents(undefined, 50, undefined, 'not-a-date'),
      /Invalid endDate/,
    );
  });

  it('parseExpandedMultistatus extracts every VEVENT occurrence from a REPORT response', () => {
    const xml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/trash.ics</D:href>
    <D:propstat>
      <D:prop>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:trash@fm
SUMMARY:Trash day
DTSTART;VALUE=DATE:20260508
DTEND;VALUE=DATE:20260509
RECURRENCE-ID;VALUE=DATE:20260508
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
    </D:propstat>
  </D:response>
  <D:response>
    <D:href>/cal/personal/bunge.ics</D:href>
    <D:propstat>
      <D:prop>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:bunge@fm
SUMMARY:David Bunge Meeting
DTSTART:20260507T130000Z
DTEND:20260507T140000Z
RECURRENCE-ID:20260507T130000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
    </D:propstat>
  </D:response>
</D:multistatus>`;

    const events = parseExpandedMultistatus(xml, 'https://caldav.example/cal/personal/');
    assert.equal(events.length, 2);
    const titles = events.map(e => e.title).sort();
    assert.deepEqual(titles, ['David Bunge Meeting', 'Trash day']);
    const trash = events.find(e => e.title === 'Trash day')!;
    // All-day events pass through with date string unchanged (no timed
    // localization), and the localizer attaches `timezone` for shape
    // consistency but never sets startUtc/endUtc on date-only values.
    assert.equal(trash.start, '2026-05-08');
    assert.equal(trash.startUtc, undefined);
    assert.ok(trash.id.includes('_'), 'recurrence-id should be appended to id');
    const bunge = events.find(e => e.title === 'David Bunge Meeting')!;
    // Bunge: 13:00Z localized to America/New_York = 09:00 EDT (UTC-4 in May).
    // Same instant, just a different presentation; UTC original preserved.
    assert.equal(
      Date.parse(bunge.start as string),
      Date.parse('2026-05-07T13:00:00Z'),
      `bunge.start should represent 13:00Z, got ${bunge.start}`,
    );
    assert.match(bunge.start as string, /-04:00$/);
    assert.equal(bunge.startUtc, '2026-05-07T13:00:00Z');
    assert.equal(bunge.timezone, 'America/New_York');
  });

  it('queries calendars in parallel and merges before slicing (no starvation)', async () => {
    // Why: sequential iteration with `if (allEvents.length >= limit) break`
    // starves later calendars — a volume-heavy calendar fills the budget and
    // a later calendar's events (e.g. Thursday "Bunge" meeting) never get
    // queried. This test asserts that:
    //   1. ALL calendars are queried (no early break),
    //   2. Events from a later calendar are NOT lost when an earlier calendar
    //      returns more than `limit` events,
    //   3. The final merged+sorted result respects the global limit.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const heavyCal = { displayName: 'TRT', url: 'https://caldav.example/cal/trt/' };
    const lightCal = { displayName: 'Calendar', url: 'https://caldav.example/cal/personal/' };
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [heavyCal, lightCal]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    // Build a multistatus body with N VEVENTs for the given calendar URL
    function buildMultistatus(calendarUrl: string, count: number, prefix: string): string {
      const responses: string[] = [];
      for (let i = 0; i < count; i++) {
        // Stagger times so sort is deterministic. Heavy events at hour 09,
        // light events at hour 13 (Thursday meeting time).
        const isHeavy = prefix === 'trt';
        const hour = isHeavy ? '09' : '13';
        const day = String(10 + (i % 28)).padStart(2, '0');
        responses.push(
          `<D:response><D:href>${calendarUrl}${prefix}-${i}.ics</D:href><D:propstat><D:prop>` +
            `<C:calendar-data><![CDATA[BEGIN:VCALENDAR\nBEGIN:VEVENT\nUID:${prefix}-${i}@fm\n` +
            `DTSTART:202605${day}T${hour}0000Z\nDTEND:202605${day}T${hour}3000Z\n` +
            `SUMMARY:${prefix === 'trt' ? 'Supplement' : 'Bunge'} ${i}\n` +
            `RECURRENCE-ID:202605${day}T${hour}0000Z\nEND:VEVENT\nEND:VCALENDAR]]></C:calendar-data>` +
            `</D:prop></D:propstat></D:response>`,
        );
      }
      return `<?xml version="1.0" encoding="utf-8"?><D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">${responses.join('')}</D:multistatus>`;
    }

    const originalFetch = globalThis.fetch;
    const fetchedUrls: string[] = [];
    (globalThis as any).fetch = async (url: string, _init: any) => {
      fetchedUrls.push(url);
      let body: string;
      if (url.includes('/cal/trt/')) {
        body = buildMultistatus(url, 100, 'trt');
      } else if (url.includes('/cal/personal/')) {
        body = buildMultistatus(url, 5, 'bunge');
      } else {
        body = '<?xml version="1.0"?><D:multistatus xmlns:D="DAV:"></D:multistatus>';
      }
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => body,
      };
    };

    try {
      const limit = 50;
      const events = await client.getCalendarEvents(
        undefined,
        limit,
        '2026-05-01T00:00:00Z',
        '2026-05-31T00:00:00Z',
      );

      // Assert ALL calendars were queried (proves no starvation / early-exit).
      assert.equal(fetchedUrls.length, 2, 'both calendars must be queried');
      assert.ok(
        fetchedUrls.some(u => u.includes('/cal/trt/')),
        'TRT calendar must be queried',
      );
      assert.ok(
        fetchedUrls.some(u => u.includes('/cal/personal/')),
        'personal Calendar must be queried (where Bunge meeting lives)',
      );

      // Assert the result is sliced to limit (post-merge, post-sort).
      assert.equal(events.length, limit);

      // Assert all 5 Bunge events from the lighter calendar made it into the
      // candidate pool — they must not be lost because the heavy calendar
      // returned 100 events. Since events are sorted by start ASC, the first
      // events should be the May 10 ones; both calendars have entries on
      // May 10 (heavy at 09:00, light at 13:00), so the Bunge entries
      // *can* land within the top-50 slice. Verify that at least one Bunge
      // event is present in the limit-sliced result.
      const bungeInResult = events.filter(e => e.title.startsWith('Bunge'));
      assert.ok(
        bungeInResult.length > 0,
        `expected at least one Bunge event in top-${limit} result, got: ${events.map(e => e.title).slice(0, 5).join(', ')}...`,
      );
    } finally {
      (globalThis as any).fetch = originalFetch;
    }
  });

  it('continues when one calendar fails and surfaces the other calendar events', async () => {
    // Why: a transient network error or 4xx/5xx on one calendar must not
    // tank the whole broad-query response. Promise.allSettled means failed
    // calendars are logged and skipped while successful ones still return.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const goodCal = { displayName: 'Good', url: 'https://caldav.example/cal/good/' };
    const badCal = { displayName: 'Bad', url: 'https://caldav.example/cal/bad/' };
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [goodCal, badCal]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    const goodXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/good/a.ics</D:href>
    <D:propstat><D:prop>
      <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:good@fm
DTSTART:20260507T100000Z
DTEND:20260507T110000Z
SUMMARY:Survives bad-calendar failure
RECURRENCE-ID:20260507T100000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
    </D:prop></D:propstat>
  </D:response>
</D:multistatus>`;

    const originalFetch = globalThis.fetch;
    (globalThis as any).fetch = async (url: string, _init: any) => {
      if (url.includes('/cal/bad/')) {
        return {
          ok: false,
          status: 503,
          statusText: 'Service Unavailable',
          text: async () => 'upstream down',
        };
      }
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => goodXml,
      };
    };

    // Suppress expected error log so test output stays clean.
    const originalErr = console.error;
    console.error = () => {};
    try {
      const events = await client.getCalendarEvents(
        undefined,
        50,
        '2026-05-01T00:00:00Z',
        '2026-05-31T00:00:00Z',
      );
      assert.equal(events.length, 1);
      assert.equal(events[0].title, 'Survives bad-calendar failure');
    } finally {
      console.error = originalErr;
      (globalThis as any).fetch = originalFetch;
    }
  });

  it('does not pass timeRange when no dates provided', async () => {
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);
    await client.getCalendarEvents(undefined, 50);

    const callArgs = mockDAVClient.fetchCalendarObjects.mock.calls[0].arguments[0];
    assert.equal(callArgs.timeRange, undefined);
  });
});

/**
 * Why: The Bunge symptom motivated the TZ-aware refactor. The agent passes
 * naive datetimes ("2026-05-07T00:00:00") to mean "Thursday morning local";
 * before this work they were silently parsed as UTC and missed the user's
 * 09:00 EDT (= 13:00 UTC) recurring meeting. These four tests pin down the
 * resolution rules and the regression so a future refactor can't reintroduce
 * the bug.
 */
describe('Timezone-aware datetime resolution', () => {
  // Each test sets its own FASTMAIL_TIMEZONE and resets the module's cache
  // so the resolved zone is fresh per test. Save and restore env to avoid
  // bleeding state into unrelated tests.
  let savedFmTz: string | undefined;
  let savedTz: string | undefined;
  before(() => {
    savedFmTz = process.env.FASTMAIL_TIMEZONE;
    savedTz = process.env.TZ;
  });
  after(() => {
    if (savedFmTz === undefined) delete process.env.FASTMAIL_TIMEZONE;
    else process.env.FASTMAIL_TIMEZONE = savedFmTz;
    if (savedTz === undefined) delete process.env.TZ;
    else process.env.TZ = savedTz;
    __resetTimezoneCacheForTests();
  });

  it('(a) resolves naive datetime in America/New_York to UTC (EDT in May = UTC-4)', () => {
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();
    // 2026-05-07 falls in EDT (DST is in effect in May), so 00:00 local = 04:00Z.
    // luxon emits ISO with millisecond precision; assert the instant equality
    // rather than the exact string so a future formatting tweak doesn't break us.
    const resolved = resolveDateInput('2026-05-07T00:00:00');
    assert.equal(
      Date.parse(resolved),
      Date.parse('2026-05-07T04:00:00Z'),
      `naive 2026-05-07T00:00:00 in America/New_York should resolve to 2026-05-07T04:00:00Z, got ${resolved}`,
    );
  });

  it('(b) preserves Z-suffixed input regardless of default TZ', () => {
    // Set default TZ to something exotic to prove the explicit Z wins.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();
    const resolved = resolveDateInput('2026-05-07T00:00:00Z');
    assert.equal(
      Date.parse(resolved),
      Date.parse('2026-05-07T00:00:00Z'),
      `Z-suffixed input must be preserved as the same UTC instant, got ${resolved}`,
    );
  });

  it('(c) honors explicit -05:00 offset regardless of default TZ', () => {
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();
    // -05:00 means the wall-clock 00:00 is 5 hours behind UTC → UTC is 05:00Z.
    const resolved = resolveDateInput('2026-05-07T00:00:00-05:00');
    assert.equal(
      Date.parse(resolved),
      Date.parse('2026-05-07T05:00:00Z'),
      `offset -05:00 input must resolve to 2026-05-07T05:00:00Z, got ${resolved}`,
    );
  });

  it('(d) Bunge regression: 13:00Z event IS included in naive 00:00→12:00 query with America/New_York default TZ', async () => {
    // The acceptance criterion. Naive 00:00 → 12:00 in EDT = 04:00Z → 16:00Z,
    // and the recurring Bunge meeting at 13:00Z (= 09:00 EDT) falls inside.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const personalCal = { displayName: 'Personal', url: 'https://caldav.example/cal/personal/' };
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [personalCal]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    // Capture the REPORT body so we can assert the bounds the server received.
    const fetchCalls: any[] = [];
    const expandedXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/bunge.ics</D:href>
    <D:propstat>
      <D:prop>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:bunge@fm
SUMMARY:David Bunge Meeting
DTSTART:20260507T130000Z
DTEND:20260507T140000Z
RECURRENCE-ID:20260507T130000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
    </D:propstat>
  </D:response>
</D:multistatus>`;
    const originalFetch = globalThis.fetch;
    (globalThis as any).fetch = async (url: string, init: any) => {
      fetchCalls.push({ url, init });
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => expandedXml,
      };
    };

    try {
      // Pass the agent's "Thursday morning" exactly as it would today: naive,
      // no Z, no offset. With the TZ-aware resolution wired in, this should
      // become a 04:00Z → 16:00Z window.
      const events = await client.getCalendarEvents(
        undefined,
        50,
        '2026-05-07T00:00:00',
        '2026-05-07T12:00:00',
      );

      // Verify the server received the EDT-shifted UTC bounds, not the naive
      // strings interpreted as UTC.
      assert.equal(fetchCalls.length, 1, 'expected exactly one REPORT request');
      const body = fetchCalls[0].init.body as string;
      assert.match(
        body,
        /<C:expand\s+start="20260507T040000Z"\s+end="20260507T160000Z"\s*\/>/,
        `expected expand bounds 04:00Z → 16:00Z (EDT-shifted from naive 00:00→12:00), got body: ${body}`,
      );

      // The event at 13:00Z falls inside the EDT-shifted window and must be
      // returned. Before the TZ fix, the window was 00:00Z → 12:00Z and the
      // event was excluded — that's the Bunge bug we're regression-testing.
      assert.equal(
        events.length,
        1,
        `Bunge event at 13:00Z must be included in the EDT-shifted morning window, got events: ${JSON.stringify(events)}`,
      );
      assert.equal(events[0].title, 'David Bunge Meeting');
    } finally {
      (globalThis as any).fetch = originalFetch;
      __resetTimezoneCacheForTests();
    }
  });
});

/**
 * Why: After 940aa79 added TZ-aware naive parsing, the agent kept attaching `Z`
 * to its query window edges anyway — so EDT mornings still got read as UTC and
 * the Bunge meeting still got dropped. The user's directive was "make the
 * SERVER force local TZ; ignore whatever Z the agent slaps on". These tests
 * pin down the FASTMAIL_FORCE_LOCAL_TZ knob and the regression that motivated
 * it: with force-local on, a Z-suffixed naive 00:00→12:00 window must be
 * treated as 00:00→12:00 EDT (= 04:00Z→16:00Z) and include the 13:00Z Bunge
 * event.
 */
describe('Force-local timezone mode (FASTMAIL_FORCE_LOCAL_TZ)', () => {
  // Same env save/restore discipline as the TZ-aware block above. Each test
  // explicitly sets/unsets FASTMAIL_FORCE_LOCAL_TZ + FASTMAIL_TIMEZONE and
  // resets the resolver cache so zones are fresh per test.
  let savedFmTz: string | undefined;
  let savedTz: string | undefined;
  let savedForceLocal: string | undefined;
  before(() => {
    savedFmTz = process.env.FASTMAIL_TIMEZONE;
    savedTz = process.env.TZ;
    savedForceLocal = process.env.FASTMAIL_FORCE_LOCAL_TZ;
  });
  after(() => {
    if (savedFmTz === undefined) delete process.env.FASTMAIL_TIMEZONE;
    else process.env.FASTMAIL_TIMEZONE = savedFmTz;
    if (savedTz === undefined) delete process.env.TZ;
    else process.env.TZ = savedTz;
    if (savedForceLocal === undefined) delete process.env.FASTMAIL_FORCE_LOCAL_TZ;
    else process.env.FASTMAIL_FORCE_LOCAL_TZ = savedForceLocal;
    __resetTimezoneCacheForTests();
  });

  it('isForceLocalTimezone() recognizes truthy values and rejects everything else', () => {
    // Truthy values
    for (const truthy of ['1', 'true', 'TRUE', 'True', 'yes', 'YES', 'on', 'ON']) {
      process.env.FASTMAIL_FORCE_LOCAL_TZ = truthy;
      assert.equal(
        isForceLocalTimezone(),
        true,
        `expected ${JSON.stringify(truthy)} to be truthy`,
      );
    }
    // Falsy values
    for (const falsy of ['0', 'false', 'no', 'off', '', 'maybe', 'foo']) {
      process.env.FASTMAIL_FORCE_LOCAL_TZ = falsy;
      assert.equal(
        isForceLocalTimezone(),
        false,
        `expected ${JSON.stringify(falsy)} to be falsy`,
      );
    }
    // Unset
    delete process.env.FASTMAIL_FORCE_LOCAL_TZ;
    assert.equal(isForceLocalTimezone(), false, 'unset env var must be falsy');
  });

  it('(a) with force-local ON: Z-suffixed input is stripped and re-localized in America/New_York', () => {
    // 2026-05-07 is in EDT (UTC-4). Naive 00:00 in EDT = 04:00Z. With force-local
    // ON, the agent's `Z` suffix should be discarded and the wall-clock time
    // re-interpreted in America/New_York → 04:00Z.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    process.env.FASTMAIL_FORCE_LOCAL_TZ = 'true';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const resolved = resolveDateInput('2026-05-07T00:00:00Z');
    assert.equal(
      Date.parse(resolved),
      Date.parse('2026-05-07T04:00:00Z'),
      `force-local should strip Z and re-localize: 2026-05-07T00:00:00Z (with EDT default) → 2026-05-07T04:00:00Z, got ${resolved}`,
    );
  });

  it('(b) with force-local OFF (default): Z-suffixed input is preserved as the same UTC instant', () => {
    // Existing behavior — when the user hasn't opted into force-local, an
    // explicit `Z` is still honored. This is the regression guard against
    // accidentally enabling the override globally.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.FASTMAIL_FORCE_LOCAL_TZ;
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const resolved = resolveDateInput('2026-05-07T00:00:00Z');
    assert.equal(
      Date.parse(resolved),
      Date.parse('2026-05-07T00:00:00Z'),
      `force-local OFF should preserve Z: 2026-05-07T00:00:00Z → 2026-05-07T00:00:00Z, got ${resolved}`,
    );
  });

  it('(c) with force-local ON: explicit -05:00 offset is stripped and re-localized in America/New_York', () => {
    // The agent might also try `-05:00` instead of `Z`. Force-local should
    // strip that too and re-interpret the wall-clock 00:00 as EDT (UTC-4),
    // producing 04:00Z — NOT the 05:00Z the offset would naively yield.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    process.env.FASTMAIL_FORCE_LOCAL_TZ = 'true';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const resolved = resolveDateInput('2026-05-07T00:00:00-05:00');
    assert.equal(
      Date.parse(resolved),
      Date.parse('2026-05-07T04:00:00Z'),
      `force-local should strip -05:00 offset and re-localize as EDT: 2026-05-07T00:00:00-05:00 → 2026-05-07T04:00:00Z, got ${resolved}`,
    );
  });

  it('(d) Bunge regression with force-local ON + Z-suffixed input: 13:00Z event IS included', async () => {
    // The acceptance criterion that motivated this whole knob. The agent
    // passes Z-suffixed naive datetimes (`2026-05-07T00:00:00Z` /
    // `2026-05-07T12:00:00Z`) meaning "Thursday morning". With force-local
    // ON the server strips the Z, treats them as EDT, and queries the server
    // with 04:00Z → 16:00Z. The 13:00Z Bunge meeting (= 09:00 EDT) falls
    // inside that window and must be returned.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    process.env.FASTMAIL_FORCE_LOCAL_TZ = 'true';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const personalCal = { displayName: 'Personal', url: 'https://caldav.example/cal/personal/' };
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [personalCal]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    const fetchCalls: any[] = [];
    const expandedXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/bunge.ics</D:href>
    <D:propstat>
      <D:prop>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:bunge@fm
SUMMARY:David Bunge Meeting
DTSTART:20260507T130000Z
DTEND:20260507T140000Z
RECURRENCE-ID:20260507T130000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
    </D:propstat>
  </D:response>
</D:multistatus>`;
    const originalFetch = globalThis.fetch;
    (globalThis as any).fetch = async (url: string, init: any) => {
      fetchCalls.push({ url, init });
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => expandedXml,
      };
    };

    // Suppress the force-local-tz debug log so test output stays clean.
    const originalErr = console.error;
    console.error = () => {};
    try {
      // Pass the agent's "Thursday morning" with the Z it actually sends.
      // BEFORE force-local: bounds become 00:00Z→12:00Z, Bunge at 13:00Z is
      // EXCLUDED. AFTER force-local: bounds become 04:00Z→16:00Z, Bunge is
      // INCLUDED. That's the regression.
      const events = await client.getCalendarEvents(
        undefined,
        50,
        '2026-05-07T00:00:00Z',
        '2026-05-07T12:00:00Z',
      );

      assert.equal(fetchCalls.length, 1, 'expected exactly one REPORT request');
      const body = fetchCalls[0].init.body as string;
      assert.match(
        body,
        /<C:expand\s+start="20260507T040000Z"\s+end="20260507T160000Z"\s*\/>/,
        `force-local should produce EDT-shifted bounds 04:00Z→16:00Z from Z-suffixed naive 00:00→12:00, got body: ${body}`,
      );

      assert.equal(
        events.length,
        1,
        `Bunge event at 13:00Z must be included in the EDT-shifted morning window when force-local is ON, got events: ${JSON.stringify(events)}`,
      );
      assert.equal(events[0].title, 'David Bunge Meeting');
    } finally {
      console.error = originalErr;
      (globalThis as any).fetch = originalFetch;
      __resetTimezoneCacheForTests();
    }
  });
});

/**
 * Why: The Bunge regression had a second face — even after the input-side TZ
 * fixes (940aa79, 404dbf7) made the QUERY window correct, the agent kept
 * dropping the event client-side because the OUTPUT was still in UTC. The
 * agent saw `2026-05-07T13:00:00Z` and reasoned "1pm UTC = afternoon",
 * filtering out a 09:00 EDT meeting that IS in the user's morning. Surfacing
 * the local-zone representation with explicit offset removes that misread.
 * These tests pin down the output-localization contract: what shape the
 * agent sees in JSON for timed events, all-day events, and the full
 * end-to-end Bunge flow.
 */
describe('Event output localization', () => {
  // Same env discipline as the input-resolution tests above. Each test sets
  // FASTMAIL_TIMEZONE explicitly and resets the resolver cache so zones are
  // fresh per test, then restores prior env.
  let savedFmTz: string | undefined;
  let savedTz: string | undefined;
  before(() => {
    savedFmTz = process.env.FASTMAIL_TIMEZONE;
    savedTz = process.env.TZ;
  });
  after(() => {
    if (savedFmTz === undefined) delete process.env.FASTMAIL_TIMEZONE;
    else process.env.FASTMAIL_TIMEZONE = savedFmTz;
    if (savedTz === undefined) delete process.env.TZ;
    else process.env.TZ = savedTz;
    __resetTimezoneCacheForTests();
  });

  it('(a) NY zone: 13:00Z event localized to 09:00-04:00 with UTC preserved', () => {
    // Acceptance: with defaultTimezone=America/New_York, an event whose
    // CalDAV dtstart is 2026-05-07T13:00:00Z is returned with
    // start="2026-05-07T09:00:00-04:00" (or millisecond-precise equivalent),
    // startUtc="2026-05-07T13:00:00Z", timezone="America/New_York".
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const localized = localizeEventTimes({
      start: '2026-05-07T13:00:00Z',
      end: '2026-05-07T14:00:00Z',
    });

    // The localized start represents the same UTC instant — same moment,
    // different presentation. Date.parse equivalence is the cleanest way to
    // assert that without coupling to luxon's millisecond formatting.
    assert.equal(
      Date.parse(localized.start as string),
      Date.parse('2026-05-07T13:00:00Z'),
      `localized start should be the same instant as 13:00Z, got ${localized.start}`,
    );
    // Explicit -04:00 offset is the human-visible signal that times are EDT.
    // EDT is UTC-4, in effect on 2026-05-07 (DST runs from 2nd Sun in March).
    assert.match(localized.start as string, /-04:00$/, `start must end with -04:00 offset, got ${localized.start}`);
    // End must also be localized.
    assert.equal(
      Date.parse(localized.end as string),
      Date.parse('2026-05-07T14:00:00Z'),
      `localized end should be the same instant as 14:00Z, got ${localized.end}`,
    );
    assert.match(localized.end as string, /-04:00$/);
    // Original UTC instants preserved verbatim — required for round-tripping
    // back to update_calendar_event without rounding shifts.
    assert.equal(localized.startUtc, '2026-05-07T13:00:00Z');
    assert.equal(localized.endUtc, '2026-05-07T14:00:00Z');
    assert.equal(localized.timezone, 'America/New_York');
  });

  it('(b) all-day event: date string preserved unchanged, no time component added', () => {
    // Why: Fastmail signals all-day events with a date-only DTSTART
    // (DTSTART;VALUE=DATE:20260507). After formatICalDate that becomes the
    // string "2026-05-07" — no T, no time. The localizer must NOT mangle
    // that into a midnight datetime; an all-day birthday is not a 00:00
    // timed event. Date-only stays date-only; only `timezone` is added so
    // the agent has consistent shape signals.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const localized = localizeEventTimes({
      start: '2026-05-07',
      end: '2026-05-08',
    });

    assert.equal(localized.start, '2026-05-07', 'all-day start must pass through unchanged');
    assert.equal(localized.end, '2026-05-08', 'all-day end must pass through unchanged');
    // No UTC original for date-only values — emitting a fake `00:00:00Z`
    // would mislead consumers about the event's semantics (all-day vs timed).
    assert.equal(localized.startUtc, undefined, 'all-day start must not have a startUtc field');
    assert.equal(localized.endUtc, undefined, 'all-day end must not have an endUtc field');
    // Timezone field is still set so the agent gets a consistent shape.
    assert.equal(localized.timezone, 'America/New_York');
  });

  it('(c) Bunge end-to-end: getCalendarEvents on broad mock returns Bunge with 09:00-04:00, not 13:00Z', async () => {
    // The whole-system regression: a naive 00:00 → 12:00 query against a
    // mocked CalDAV serving Bunge at 13:00Z (timed) AND a "Trash day"
    // all-day supplement. After the localizer, the JSON the MCP layer hands
    // to the agent must show Bunge with start="...09:00:00...-04:00" and the
    // all-day event still as "2026-05-08". This is the shape that prevents
    // the agent's afternoon/morning misread.
    process.env.FASTMAIL_TIMEZONE = 'America/New_York';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const personalCal = { displayName: 'Personal', url: 'https://caldav.example/cal/personal/' };
    const supplementCal = { displayName: 'Supplements', url: 'https://caldav.example/cal/supp/' };
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [personalCal, supplementCal]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    const personalXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/personal/bunge.ics</D:href>
    <D:propstat>
      <D:prop>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:bunge@fm
SUMMARY:David Bunge Meeting
DTSTART:20260507T130000Z
DTEND:20260507T140000Z
RECURRENCE-ID:20260507T130000Z
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
    </D:propstat>
  </D:response>
</D:multistatus>`;
    const supplementXml = `<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>/cal/supp/trash.ics</D:href>
    <D:propstat>
      <D:prop>
        <C:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:trash@fm
SUMMARY:Trash day
DTSTART;VALUE=DATE:20260507
DTEND;VALUE=DATE:20260508
RECURRENCE-ID;VALUE=DATE:20260507
END:VEVENT
END:VCALENDAR]]></C:calendar-data>
      </D:prop>
    </D:propstat>
  </D:response>
</D:multistatus>`;

    const originalFetch = globalThis.fetch;
    (globalThis as any).fetch = async (url: string) => {
      const xml = url.includes('/personal/') ? personalXml : supplementXml;
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        text: async () => xml,
      };
    };

    try {
      // Agent's "Thursday morning" — naive, no Z, no offset. With NY default
      // TZ, this becomes a 04:00Z → 16:00Z window; both calendars are queried
      // in parallel and merged.
      const events = await client.getCalendarEvents(
        undefined,
        50,
        '2026-05-07T00:00:00',
        '2026-05-07T12:00:00',
      );

      const bunge = events.find(e => e.title === 'David Bunge Meeting');
      const trash = events.find(e => e.title === 'Trash day');

      // Bunge present and localized (NOT shown as 13:00Z, which is the bug).
      assert.ok(bunge, `Bunge must be in result; events=${JSON.stringify(events)}`);
      assert.match(
        bunge!.start as string,
        /09:00:00.*-04:00$/,
        `Bunge start must show 09:00:00 with -04:00 offset (NY morning), got ${bunge!.start}`,
      );
      assert.equal(bunge!.startUtc, '2026-05-07T13:00:00Z', 'Bunge startUtc must preserve the original UTC instant');
      assert.equal(bunge!.timezone, 'America/New_York');

      // All-day event still passes through unchanged — the localizer must
      // NOT corrupt date-only values into datetimes.
      assert.ok(trash, 'all-day Trash event must also be in result');
      assert.equal(trash!.start, '2026-05-07', 'all-day start must remain date-only');
      assert.equal(trash!.end, '2026-05-08', 'all-day end must remain date-only');
      assert.equal(trash!.startUtc, undefined);
    } finally {
      (globalThis as any).fetch = originalFetch;
      __resetTimezoneCacheForTests();
    }
  });

  it('(d) UTC zone: 13:00Z passes through with no offset shift (localizer is configurable)', () => {
    // Sanity check that the localizer is not hardcoded to EDT. With
    // defaultTimezone=UTC, the same 13:00Z input must come out as a UTC
    // instant — formatted with `Z` suffix, no offset shift. This is the
    // baseline that proves the zone parameter actually drives behavior.
    process.env.FASTMAIL_TIMEZONE = 'UTC';
    delete process.env.TZ;
    __resetTimezoneCacheForTests();

    const localized = localizeEventTimes({
      start: '2026-05-07T13:00:00Z',
      end: '2026-05-07T14:00:00Z',
    });

    // Same instant, presented in UTC.
    assert.equal(
      Date.parse(localized.start as string),
      Date.parse('2026-05-07T13:00:00Z'),
      `start should still represent 13:00Z, got ${localized.start}`,
    );
    // luxon emits Z (not +00:00) for the UTC zone; allow either since both
    // are valid ISO 8601 representations of UTC.
    assert.match(
      localized.start as string,
      /(Z|\+00:00)$/,
      `UTC zone output must end with Z or +00:00, got ${localized.start}`,
    );
    assert.equal(localized.startUtc, '2026-05-07T13:00:00Z');
    assert.equal(localized.endUtc, '2026-05-07T14:00:00Z');
    assert.equal(localized.timezone, 'UTC');
  });
});
