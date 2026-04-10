#!/usr/bin/env node
/**
 * Why: Attaches a birthday photo to a specific CalDAV event by embedding it as base64 ATTACH.
 * What: GETs the current ICS, inserts a folded base64 ATTACH line, PUTs it back.
 * Test: Run and assert HTTP 204 (or 200/201) from the PUT response.
 */

import { readFileSync } from 'fs';

const BASE_URL = 'https://caldav.fastmail.com/dav/calendars/user/davidgutowsky@fastmail.com';
const CALENDAR_ID = '4c646201-472c-4377-b4c8-10f4455c6ecf';
const EVENT_ID = '1775780794196-lk8xsujiki@fastmail-mcp';
const PHOTO_PATH = '/home/david/.openclaw/media/inbound/file_0---127f05e7-cd9c-4688-9d62-e33c7b9efdb1.jpg';

const EVENT_URL = `${BASE_URL}/${CALENDAR_ID}/${EVENT_ID}.ics`;
const AUTH = Buffer.from('davidgutowsky@fastmail.com:7k662x973b4k5u7t').toString('base64');

async function main() {
  // Step 1: GET the current ICS
  const getResp = await fetch(EVENT_URL, {
    method: 'GET',
    headers: { Authorization: `Basic ${AUTH}` },
  });
  if (!getResp.ok) {
    throw new Error(`GET failed: ${getResp.status} ${getResp.statusText}`);
  }
  let ics = await getResp.text();
  console.log(`GET status: ${getResp.status}`);

  // Step 2: Read and base64-encode the photo
  const photoData = readFileSync(PHOTO_PATH);
  const b64 = photoData.toString('base64');
  // Fold at 75 chars per RFC 5545
  const folded = b64.match(/.{1,75}/g)?.join('\r\n ') ?? b64;
  const attachLine = `ATTACH;ENCODING=BASE64;FMTTYPE=image/jpeg:\r\n ${folded}`;

  // Step 3: Remove any existing ATTACH lines, then insert before END:VEVENT
  ics = ics.replace(/^ATTACH[^\r\n]*(\r?\n[ \t][^\r\n]*)*/gm, '');
  ics = ics.replace(/(\r?\n){2,}/g, '\r\n');
  ics = ics.replace(/END:VEVENT/, `${attachLine}\r\nEND:VEVENT`);

  // Step 4: PUT the updated ICS back
  const putResp = await fetch(EVENT_URL, {
    method: 'PUT',
    headers: {
      Authorization: `Basic ${AUTH}`,
      'Content-Type': 'text/calendar; charset=utf-8',
    },
    body: ics,
  });
  console.log(`PUT status: ${putResp.status} ${putResp.statusText}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
