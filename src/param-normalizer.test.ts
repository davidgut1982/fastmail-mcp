import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { normalizeToolArgs, __test } from './param-normalizer.js';

// Mirror of the inputSchema we use for create_calendar_event in src/index.ts.
// Pinned here so test cases are self-contained and survive future schema
// edits without false-passing.
const createCalendarEventSchema = {
  type: 'object',
  properties: {
    calendarId: { type: 'string' },
    title: { type: 'string' },
    description: { type: 'string' },
    start: { type: 'string' },
    end: { type: 'string' },
    location: { type: 'string' },
    participants: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          email: { type: 'string' },
          name: { type: 'string' },
        },
      },
    },
  },
  required: ['calendarId', 'title', 'start', 'end'],
};

// Synthetic schema with a nested object whose property is camelCase. Used to
// confirm we recurse and rename inside nested objects, not just at the root.
const nestedSchema = {
  type: 'object',
  properties: {
    filter: {
      type: 'object',
      properties: {
        afterDate: { type: 'string' },
        beforeDate: { type: 'string' },
      },
    },
  },
};

describe('snakeToCamel', () => {
  it('converts single-underscore snake_case', () => {
    assert.equal(__test.snakeToCamel('calendar_id'), 'calendarId');
  });
  it('converts multi-underscore snake_case', () => {
    assert.equal(__test.snakeToCamel('start_date_time'), 'startDateTime');
  });
  it('passes through camelCase unchanged', () => {
    assert.equal(__test.snakeToCamel('calendarId'), 'calendarId');
  });
  it('passes through single words unchanged', () => {
    assert.equal(__test.snakeToCamel('title'), 'title');
  });
});

describe('normalizeToolArgs', () => {
  it('a. renames calendar_id → calendarId for create_calendar_event', () => {
    const input = {
      calendar_id: 'cal-1',
      title: 'Meeting',
      start: '2026-05-07T09:00:00',
      end: '2026-05-07T10:00:00',
    };
    const result = normalizeToolArgs(input, createCalendarEventSchema, 'create_calendar_event');
    assert.equal(result.changed, true);
    assert.equal(result.args.calendarId, 'cal-1');
    assert.equal((result.args as any).calendar_id, undefined);
    assert.equal(result.args.title, 'Meeting');
    assert.deepEqual(
      result.renames.map((r) => `${r.from}->${r.to}`),
      ['calendar_id->calendarId']
    );
  });

  it('b. passes camelCase calendarId through unchanged (no rename, no log entry)', () => {
    const input = {
      calendarId: 'cal-2',
      title: 'Meeting',
      start: '2026-05-07T09:00:00',
      end: '2026-05-07T10:00:00',
    };
    const result = normalizeToolArgs(input, createCalendarEventSchema, 'create_calendar_event');
    assert.equal(result.changed, false);
    assert.equal(result.renames.length, 0);
    assert.equal(result.args.calendarId, 'cal-2');
  });

  it('c. mixed args: keeps camelCase, renames snake_case', () => {
    const input = {
      calendarId: 'cal-3',
      title: 'Mixed',
      // Also throw in a snake field that doesn't match any known camelCase
      // and confirm it survives untouched.
      start_date: 'unused-but-not-renamed-because-no-startDate-key', // not in schema
      start: '2026-05-07T09:00:00',
      end: '2026-05-07T10:00:00',
    };
    const result = normalizeToolArgs(input, createCalendarEventSchema, 'create_calendar_event');
    // calendar_id wasn't sent so no rename; start_date doesn't have a
    // matching camelCase key in this schema, so it passes through.
    assert.equal(result.args.calendarId, 'cal-3');
    assert.equal((result.args as any).start_date, 'unused-but-not-renamed-because-no-startDate-key');
    assert.equal(result.changed, false);

    // Now exercise the actual mixed case from the task description.
    const result2 = normalizeToolArgs(
      { calendarId: 'x', start_date: 'y' },
      // Synthesize a schema where startDate IS a known camelCase key.
      {
        type: 'object',
        properties: {
          calendarId: { type: 'string' },
          startDate: { type: 'string' },
        },
      },
      'synthetic'
    );
    assert.equal(result2.changed, true);
    assert.equal(result2.args.calendarId, 'x');
    assert.equal(result2.args.startDate, 'y');
    assert.equal((result2.args as any).start_date, undefined);
  });

  it('d. nested object: filter.after_date → filter.afterDate', () => {
    const input = {
      filter: {
        after_date: '2026-01-01',
        before_date: '2026-12-31',
      },
    };
    const result = normalizeToolArgs(input, nestedSchema, 'list_things');
    assert.equal(result.changed, true);
    assert.equal(result.args.filter.afterDate, '2026-01-01');
    assert.equal(result.args.filter.beforeDate, '2026-12-31');
    assert.equal(result.args.filter.after_date, undefined);
  });

  it('e. unknown snake_case keys are NOT renamed (avoids false positives)', () => {
    const input = {
      calendarId: 'cal-4',
      title: 'X',
      start: '2026-05-07T09:00:00',
      end: '2026-05-07T10:00:00',
      random_unknown_key: 'should-survive',
    };
    const result = normalizeToolArgs(input, createCalendarEventSchema, 'create_calendar_event');
    assert.equal(result.changed, false);
    assert.equal((result.args as any).random_unknown_key, 'should-survive');
  });

  it('idempotent: running twice yields the same output', () => {
    const input = {
      calendar_id: 'cal-5',
      title: 'Y',
      start: '2026-05-07T09:00:00',
      end: '2026-05-07T10:00:00',
    };
    const once = normalizeToolArgs(input, createCalendarEventSchema, 'create_calendar_event');
    const twice = normalizeToolArgs(once.args, createCalendarEventSchema, 'create_calendar_event');
    assert.equal(twice.changed, false);
    assert.deepEqual(twice.args, once.args);
  });

  it('handles array of objects: participant entries get nested normalization', () => {
    // create_calendar_event.participants items have email/name (no snake_case
    // keys to test directly), so synthesize one with a snake_case nested key.
    const schema = {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              displayName: { type: 'string' },
              emailAddress: { type: 'string' },
            },
          },
        },
      },
    };
    const input = {
      items: [
        { display_name: 'Alice', email_address: 'alice@example.com' },
        { displayName: 'Bob', emailAddress: 'bob@example.com' },
      ],
    };
    const result = normalizeToolArgs(input, schema, 'syn');
    assert.equal(result.changed, true);
    assert.equal(result.args.items[0].displayName, 'Alice');
    assert.equal(result.args.items[0].emailAddress, 'alice@example.com');
    assert.equal(result.args.items[1].displayName, 'Bob');
    assert.equal(result.args.items[1].emailAddress, 'bob@example.com');
  });

  it('handles missing/empty inputs gracefully', () => {
    assert.deepEqual(
      normalizeToolArgs(undefined, createCalendarEventSchema, 't').args,
      {}
    );
    assert.deepEqual(
      normalizeToolArgs(null as any, createCalendarEventSchema, 't').args,
      {}
    );
    assert.deepEqual(
      normalizeToolArgs({}, createCalendarEventSchema, 't').args,
      {}
    );
    // Empty schema (no properties) → pass through unchanged.
    const r = normalizeToolArgs({ foo_bar: 1 }, { type: 'object' }, 't');
    assert.equal(r.changed, false);
    assert.deepEqual(r.args, { foo_bar: 1 });
  });
});
