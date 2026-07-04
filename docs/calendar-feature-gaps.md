# fastmail-mcp Calendar Feature Gap Analysis

**Date:** 2026-07-04
**Scope:** CalDAV + JMAP calendar paths in fastmail-mcp v1.9.1

## Baseline (what IS supported)

| Property | Read | Write | Notes |
|----------|------|-------|-------|
| SUMMARY | ✅ | ✅ | Event title |
| DESCRIPTION | ✅ | ✅ | Event description |
| DTSTART | ✅ | ✅ | Start time (TZ-aware) |
| DTEND | ✅ | ✅ | End time (TZ-aware) |
| LOCATION | ✅ | ✅ | Event location |
| UID | ✅ | ✅ | Auto-generated |
| DTSTAMP | ✅ | ✅ | Auto-generated |
| ATTENDEE | ✅ | ✅ | Email + name + RSVP |
| ORGANIZER | ✗ | ✅ | Hard-coded identity |
| ATTACH | ✗ | ✅ | Base64 attachments via update |
| RECURRENCE-ID | ✅ | n/a | Expanded occurrences only |

**Read-path parser** (`parseCalendarObject` / `parseExpandedMultistatus` in `caldav-client.ts`) silently drops every property not listed above. Any RRULE, VALARM, STATUS, etc. present in the server's ICS is invisible to the agent.

---

## HIGH severity (commonly needed)

### 1. VALARM / Alarms
- **RFC 5545:** §3.6.6
- **Missing:** `caldav-client.ts:891-908` (create), `caldav-client.ts:733-841` (update), `index.ts:568-652` (schemas)
- **Needed:** `alarms?: Array<{ trigger: string; action?: 'DISPLAY'|'EMAIL'; description?: string }>` param; emit `BEGIN:VALARM…END:VALARM` inside VEVENT

### 2. RRULE / Recurrence rule (create/update)
- **RFC 5545:** §3.3.10, §3.6.1
- **Missing:** create/update emit no RRULE; no schema param
- **Needed:** `recurrence?: { freq: 'DAILY'|'WEEKLY'|'MONTHLY'|'YEARLY'; interval?: number; until?: string; count?: number; byDay?: string[]; byMonthDay?: number[] }` → `RRULE:FREQ=WEEKLY;BYDAY=TH;COUNT=10`

### 3. All-day event creation (write path)
- **RFC 5545:** §3.3.4 (VALUE=DATE)
- **Missing:** `caldav-client.ts:877-878` always emits `YYYYMMDDTHHMMSSZ`; no `allDay` param in create/update
- **Note:** Read path correctly handles `DTSTART;VALUE=DATE`; gap is write-only
- **Needed:** `allDay?: boolean`; when true emit `DTSTART;VALUE=DATE:YYYYMMDD`

### 4. STATUS (TENTATIVE / CONFIRMED / CANCELLED)
- **RFC 5545:** §3.8.1.11
- **Missing:** Not parsed, not emitted
- **Needed:** `status?: 'TENTATIVE'|'CONFIRMED'|'CANCELLED'` in schema + ICS; parse on read

### 5. ORGANIZER parsing + configurable identity
- **RFC 5545:** §3.8.4.3
- **Missing:** Not parsed on read; hard-coded `davidgutowsky@fastmail.com` on write (`caldav-client.ts:798,904`; `contacts-calendar.ts:246-251`)
- **Needed:** Parse ORGANIZER into `{ email, name }`; make configurable via env or param

---

## MEDIUM severity (useful)

### 6. TRANSP / Free-Busy transparency
- **RFC 5545:** §3.8.2.7
- **Needed:** `freeBusy?: 'busy'|'free'` → `TRANSP:OPAQUE`/`TRANSP:TRANSPARENT`

### 7. CATEGORIES (tags/labels)
- **RFC 5545:** §3.8.1.2
- **Needed:** `categories?: string[]` → `CATEGORIES:Personal,Birthday`

### 8. PRIORITY
- **RFC 5545:** §3.8.1.9 (1=high … 9=low, 0=undefined)
- **Needed:** `priority?: number` → `PRIORITY:1`

### 9. CLASS / Privacy
- **RFC 5545:** §3.8.1.3
- **Needed:** `visibility?: 'public'|'private'|'confidential'` → `CLASS:PRIVATE`

### 10. EXDATE (recurrence exception dates)
- **RFC 5545:** §3.8.5.1
- **Needed:** `exdates?: string[]`; parse on read

### 11. RDATE (additional recurrence dates)
- **RFC 5545:** §3.8.5.2
- **Needed:** `rdates?: string[]`; parse on read

### 12. SEQUENCE (revision tracking)
- **RFC 5545:** §3.8.7.4
- **Missing:** Not bumped on update — calendar clients may not refresh
- **Needed:** Parse existing SEQUENCE; increment on update; default 0 on create

### 13. Recurrence on unbounded read path
- **RFC 4791:** §9.6.5
- **Missing:** No-date-bounds fallback (`caldav-client.ts:586-595`) doesn't expand recurrences
- **Needed:** Always issue expand REPORT with wide default window, or document the limitation

---

## LOW severity (edge cases)

### 14. GEO — §3.8.1.6
### 15. RESOURCES — §3.8.1.10
### 16. RELATED-TO — §3.8.4.4
### 17. URL (conference links) — §3.8.4.6
### 18. CREATED / LAST-MODIFIED — §3.8.7.1/3

---

## JMAP path parity (contacts-calendar.ts)

The JMAP path (`contacts-calendar.ts:173` `CalendarEvent/get` properties, `:230-264` create) has the **same gaps** as CalDAV plus drops attachments entirely. JSCalendar equivalents: `alerts`, `recurrenceRule`, `freeBusyStatus`, `privacy`, `priority`, `categories`.

---

## Recommended implementation order

1. **All-day write** (#3) — quick, unblocks common use case
2. **STATUS** (#4) — simple string field
3. **VALARM/alarms** (#1) — the original trigger for this audit
4. **RRULE** (#2) — most complex, highest impact
5. **ORGANIZER read + config** (#5)
6. **MEDIUM batch** (#6-9: TRANSP, CATEGORIES, PRIORITY, CLASS)
7. **SEQUENCE bumping** (#12)
8. **LOW metadata batch** (#14-18)
9. **JMAP parity** (#19) — apply same fields to JSCalendar
