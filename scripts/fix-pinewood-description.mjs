#!/usr/bin/env node
/**
 * Why: Removes an auto-generated attendee-update line from the Pinewood Derby event description.
 * What: GETs the ICS, unfolds the DESCRIPTION property, strips the unwanted segment, re-folds, PUTs it back.
 * Test: Run and assert HTTP 204; verify description no longer contains the unwanted line.
 */

const BASE_URL = 'https://caldav.fastmail.com/dav/calendars/user/davidgutowsky@fastmail.com';
const CALENDAR_ID = '4c646201-472c-4377-b4c8-10f4455c6ecf';
const EVENT_ID = '1775750315778-0lofd5vjog6i@fastmail-mcp';

const EVENT_URL = `${BASE_URL}/${CALENDAR_ID}/${EVENT_ID}.ics`;
const AUTH = Buffer.from('davidgutowsky@fastmail.com:7k662x973b4k5u7t').toString('base64');

// The unwanted segment as it appears in the ICS value (iCal-escaped)
const UNWANTED_SEGMENT = '\\n[2026-04-10] Attendees updated: rgutowsky@hotmail.com\\, genna.hibbs@gmail.com\\; test@example.com removed.';

/**
 * Why: RFC 5545 folds long lines at 75 octets. We need to unfold before string manipulation.
 * What: Joins continuation lines (lines starting with a space/tab) with the previous line.
 */
function unfoldIcs(ics) {
  return ics.replace(/\r?\n[ \t]/g, '');
}

/**
 * Why: Re-fold lines longer than 75 octets per RFC 5545 section 3.1.
 * What: Splits each property line at 75-octet boundaries, joining with CRLF + space.
 */
function foldLine(line) {
  if (line.length <= 75) return line;
  const chunks = [];
  let i = 0;
  chunks.push(line.slice(0, 75));
  i = 75;
  while (i < line.length) {
    chunks.push(line.slice(i, i + 74));
    i += 74;
  }
  return chunks.join('\r\n ');
}

function refoldIcs(ics) {
  return ics.split(/\r?\n/).map(foldLine).join('\r\n');
}

async function main() {
  // Step 1: GET the current ICS
  const getResp = await fetch(EVENT_URL, {
    method: 'GET',
    headers: { Authorization: `Basic ${AUTH}` },
  });
  if (!getResp.ok) {
    throw new Error(`GET failed: ${getResp.status} ${getResp.statusText}`);
  }
  const rawIcs = await getResp.text();
  console.log(`GET status: ${getResp.status}`);

  // Step 2: Unfold, strip the unwanted segment, re-fold
  const unfolded = unfoldIcs(rawIcs);
  if (!unfolded.includes(UNWANTED_SEGMENT)) {
    console.log('WARNING: Target segment not found after unfolding. Dumping DESCRIPTION line:');
    const m = unfolded.match(/^DESCRIPTION:.*/m);
    if (m) console.log(m[0].substring(0, 300));
    return;
  }

  const cleaned = unfolded.replace(UNWANTED_SEGMENT, '');
  const refolded = refoldIcs(cleaned);

  // Step 3: PUT the updated ICS back
  const putResp = await fetch(EVENT_URL, {
    method: 'PUT',
    headers: {
      Authorization: `Basic ${AUTH}`,
      'Content-Type': 'text/calendar; charset=utf-8',
    },
    body: refolded,
  });
  console.log(`PUT status: ${putResp.status} ${putResp.statusText}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
