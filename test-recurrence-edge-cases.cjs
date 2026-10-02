'use strict';

// Run: node test-recurrence-edge-cases.cjs (uses the methods in main.ts).
const assert = require('node:assert/strict');
const path = require('node:path');
const { loadMethods, methodNames } = require('./test-recurrence.cjs');
process.env.TZ = process.env.RECURRENCE_TEST_TZ || 'America/Toronto';
const repo = path.resolve(process.argv[2] || __dirname);
let suggested;
let cancel = false;
let manualDate;
class Modal {
  constructor(app, date, resolve) { suggested = date; this.resolve = resolve; }
  open() { this.resolve(cancel ? null : (manualDate ?? suggested)); }
}
const { plugin, moment, PatternType } = loadMethods(repo, [...methodNames,
  'correctRecurringOccurrenceDate', 'getOccurrenceDateFromGlobalId', 'parseDateFromDragText'
], Modal);
const minutes = date => (Date.parse(`${date}T00:00:00Z`) + 11644473600000) / 60000;
const pattern = {
  recurFrequency: 8204, patternType: PatternType.MonthNth, period: 1,
  patternTypeMonthNth: { dayOfWeekBits: 32, n: 5 },
  startDate: minutes('2025-10-31'), endDate: minutes('2026-09-25'),
  endType: 0x2021, firstDOW: 0, deletedInstanceDates: [], modifiedInstanceDates: []
};
const find = (rp, start, today) => plugin.findClosestOccurrence(
  { recurrencePattern: rp }, moment(start), moment(today)
)?.format('YYYY-MM-DD HH:mm');
let checks = 0;
function equal(actual, expected, label) {
  assert.equal(actual, expected, label);
  checks++;
}
// Independent calendar oracle: plain UTC Date enumeration, not the plugin's helper.
function nthDate(year, month, mask, nth) {
  const matches = [];
  for (let day = 1; day <= 31; day++) {
    const date = new Date(Date.UTC(year, month, day));
    if (date.getUTCMonth() !== month) break;
    if (mask & (1 << date.getUTCDay())) matches.push(date.toISOString().slice(0, 10));
  }
  return nth === 5 ? matches.at(-1) : matches[nth - 1];
}
async function main() {
  moment.locale('en');
  // Exercise different month lengths, leap years, periods, weekdays and masks.
  for (const year of [2024, 2025, 2026, 2028]) {
    for (const mask of [1, 2, 4, 8, 16, 32, 64, 62, 65, 127]) {
      for (const nth of [1, 2, 3, 4, 5]) {
        for (const period of [1, 2, 3, 6]) {
          const start = nthDate(year, 0, mask, nth);
          const end = nthDate(year, 6, mask, nth);
          const rp = { ...pattern, period, startDate: minutes(start), endDate: minutes(end), patternTypeMonthNth: { dayOfWeekBits: mask, n: nth } };
          equal(find(rp, `${start}T09:00:00`, `${year}-08-15T12:00:00`), `${end} 09:00`, 'Final date must be inclusive');
        }
      }
    }
  }
  moment.locale('fr');
  equal(find(pattern, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), '2026-09-25 09:00', 'Exact example in French');
  moment.locale('en');
  moment.now = () => new Date(2026, 9, 1, 12).valueOf();
  const makeMeeting = () => ({
    apptStartWhole: moment('2025-10-31T09:00:00').toISOString(),
    apptEndWhole: moment('2025-10-31T12:00:00').toISOString(),
    apptRecur: { recurrencePattern: pattern }
  });
  const meeting = makeMeeting();
  equal(await plugin.correctRecurringOccurrenceDate(meeting), true, 'Modal accepted');
  equal(suggested, '2026-09-25', 'Modal prefill');
  equal(moment(meeting.apptStartWhole).format('YYYY-MM-DD HH:mm'), '2026-09-25 09:00', 'Corrected start');
  equal(moment(meeting.apptEndWhole).format('YYYY-MM-DD HH:mm'), '2026-09-25 12:00', 'Corrected end');
  cancel = true;
  const cancelled = makeMeeting();
  const unchanged = JSON.stringify(cancelled);
  equal(await plugin.correctRecurringOccurrenceDate(cancelled), false, 'Modal cancelled');
  equal(JSON.stringify(cancelled), unchanged, 'Cancellation must not modify dates');
  cancel = false;
  const fromId = makeMeeting();
  fromId.globalAppointmentID = '0'.repeat(32) + '07ea081c'; // 2026-08-28
  equal(await plugin.correctRecurringOccurrenceDate(fromId, 'Start: Friday, September 25, 2026 9:00 AM'), true, 'Global ID accepted');
  equal(moment(fromId.apptStartWhole).format('YYYY-MM-DD HH:mm'), '2026-08-28 09:00', 'Global ID wins over drag text');
  const fromText = makeMeeting();
  equal(await plugin.correctRecurringOccurrenceDate(fromText, 'Start: Friday, September 25, 2026 9:00 AM'), true, 'Drag text accepted');
  equal(moment(fromText.apptStartWhole).format('YYYY-MM-DD HH:mm'), '2026-09-25 09:00', 'Drag text date');
  const specific = makeMeeting();
  specific.apptStartWhole = moment('2026-08-28T09:00:00').toISOString();
  specific.apptEndWhole = moment('2026-08-28T12:00:00').toISOString();
  const originalSpecific = JSON.stringify(specific);
  equal(await plugin.correctRecurringOccurrenceDate(specific), true, 'Already specific occurrence');
  equal(JSON.stringify(specific), originalSpecific, 'Already specific occurrence unchanged');
  const ics = { ...makeMeeting(), apptRecur: null };
  const originalIcs = JSON.stringify(ics);
  equal(await plugin.correctRecurringOccurrenceDate(ics), true, 'ICS bypass');
  equal(JSON.stringify(ics), originalIcs, 'ICS unchanged');
  process.stdout.write(`${checks} additional regression checks passed (${process.env.TZ}).\n`);

  // Regression assertions for the six originally reproduced issues.
  const active = { ...pattern, endDate: 0x5AE980DF, endType: 0x2023 };
  const probe = (label, actual, expected) => equal(actual, expected, label);
  const weekly = { ...pattern, recurFrequency: 8203, patternType: PatternType.Week, patternTypeWeek: { dayOfWeekBits: 32 }, endDate: 0x5AE980DF, endType: 0x2023 };
  moment.locale('fr');
  probe('Weekly Friday with French locale', find(weekly, '2025-10-31T09:00:00', '2026-09-25T12:00:00'), '2026-09-25 09:00');
  moment.locale('en');
  probe('Deleted occurrence', find({ ...active, deletedInstanceDates: [minutes('2026-09-25')] }, '2025-10-31T09:00:00', '2026-09-26T12:00:00'), '2026-08-28 09:00');
  probe('MonthEnd following April 30', find({ ...active, patternType: PatternType.MonthEnd, patternTypeMonth: { day: 31 }, startDate: minutes('2026-04-30') }, '2026-04-30T09:00:00', '2026-06-01T12:00:00'), '2026-05-31 09:00');
  probe('Day 31 series starting in February', find({ ...active, patternType: PatternType.Month, patternTypeMonth: { day: 31 }, startDate: minutes('2026-02-28') }, '2026-02-28T09:00:00', '2026-04-01T12:00:00'), '2026-03-31 09:00');
  if (process.env.TZ === 'America/Toronto') {
    probe('Second Sunday during spring clock change', find({ ...active, patternTypeMonthNth: { dayOfWeekBits: 1, n: 2 }, startDate: minutes('2026-02-08') }, '2026-02-08T09:00:00', '2026-03-08T12:00:00'), '2026-03-08 09:00');
    probe('First Sunday during autumn clock change', find({ ...active, patternTypeMonthNth: { dayOfWeekBits: 1, n: 1 }, startDate: minutes('2026-10-04') }, '2026-10-04T09:00:00', '2026-11-01T12:00:00'), '2026-11-01 09:00');
  }
  // Weekly dates must follow Outlook's firstDOW, not Moment's locale.
  for (const locale of ['en', 'fr']) {
    moment.locale(locale);
    for (let firstDOW = 0; firstDOW < 7; firstDOW++) {
      for (const period of [1, 2, 3]) {
        for (const mask of [1, 32, 42]) {
          const now = moment('2026-02-15T12:00:00');
          // Independent oracle enumerates calendar days using UTC Date.
          const startUtc = Date.UTC(2026, 0, 2);
          const firstWeekUtc = startUtc - ((5 - firstDOW + 7) % 7) * 86400000;
          const dates = [];
          for (let days = 0; days < 100; days++) {
            const date = new Date(startUtc + days * 86400000);
            if (Math.floor((date.valueOf() - firstWeekUtc) / 604800000) % period === 0 && (mask & (1 << date.getUTCDay())))
              dates.push(moment(date.toISOString().slice(0, 10) + 'T09:00:00'));
          }
          // Anchor baseTime at the first actual occurrence for this mask.
          const first = dates[0];
          const expected = dates.reduce((a, b) => Math.abs(a.diff(now)) <= Math.abs(b.diff(now)) ? a : b);
          const rp = { ...weekly, firstDOW, period, patternTypeWeek: { dayOfWeekBits: mask }, startDate: minutes(first.format('YYYY-MM-DD')) };
          equal(find(rp, first.format('YYYY-MM-DDTHH:mm:ss'), now.format('YYYY-MM-DDTHH:mm:ss')), expected.format('YYYY-MM-DD HH:mm'), `${locale}: week start ${firstDOW}, period ${period}, mask ${mask}`);
        }
      }
    }
  }
  moment.locale('en');
  const daily = { ...active, recurFrequency: 8202, patternType: PatternType.Day, period: 1440, startDate: minutes('2026-09-01') };
  const deleted = Array.from({ length: 15 }, (_, i) => minutes(`2026-09-${String(i + 10).padStart(2, '0')}`));
  equal(find({ ...daily, deletedInstanceDates: deleted }, '2026-09-01T09:00:00', '2026-09-17T12:00:00'), '2026-09-25 09:00', 'Search beyond many deleted dates');
  equal(find({ ...daily, endDate: minutes('2026-09-24'), deletedInstanceDates: deleted }, '2026-09-01T09:00:00', '2026-10-01T12:00:00'), '2026-09-09 09:00', 'Ended series with many deleted final dates');
  const allDeleted = { ...pattern, startDate: minutes('2025-10-31'), endDate: minutes('2025-10-31'), deletedInstanceDates: [minutes('2025-10-31')] };
  equal(find(allDeleted, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), undefined, 'All occurrences deleted');
  equal(find({ ...daily, endDate: minutes('2026-08-31') }, '2026-09-01T09:00:00', '2026-10-01T12:00:00'), undefined, 'Invalid range must not leak future dates');
  for (const period of [0, -1, NaN, Infinity, 1.5])
    equal(find({ ...pattern, period }, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), undefined, 'Invalid period');
  for (const type of [PatternType.HjMonth, PatternType.HjMonthNth, PatternType.HjMonthEnd])
    equal(find({ ...pattern, patternType: type }, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), undefined, 'Hijri needs explicit date');
  equal(find({ ...pattern, calendarType: 6 }, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), undefined, 'Hijri calendar marker');
  for (const rp of [allDeleted, { ...pattern, patternType: PatternType.HjMonthNth }]) {
    const file = makeMeeting();
    file.apptRecur.recurrencePattern = rp;
    manualDate = '2026-09-25';
    equal(await plugin.correctRecurringOccurrenceDate(file), true, 'Explicit date accepted');
    equal(suggested, '', 'No invented suggestion');
    equal(moment(file.apptStartWhole).format('YYYY-MM-DD HH:mm'), '2026-09-25 09:00', 'Explicit date preserved');
    manualDate = undefined;
  }
  // Exceptions replace the old slot and preserve their new time and duration.
  const moved = { originalStartTime: minutes('2026-09-25') + 540, startDateTime: minutes('2026-09-24') + 600, endDateTime: minutes('2026-09-24') + 690, overrideFlags: 0 };
  const withException = { recurrencePattern: { ...pattern, deletedInstanceDates: [minutes('2026-09-25')], modifiedInstanceDates: [minutes('2026-09-24')] }, exceptionInfo: [moved] };
  equal(plugin.findClosestOccurrence(withException, moment('2025-10-31T09:00:00'), moment('2026-10-01T12:00:00'))?.format('YYYY-MM-DD HH:mm'), '2026-09-24 10:00', 'Moved occurrence nearest to today');
  for (const mode of ['modal', 'id', 'text']) {
    const file = makeMeeting();
    file.apptRecur = withException;
    if (mode === 'id') file.globalAppointmentID = '0'.repeat(32) + '07ea0919';
    const text = mode === 'text' ? 'Start: Thursday, September 24, 2026 10:00 AM' : '';
    equal(await plugin.correctRecurringOccurrenceDate(file, text), true, `Exception via ${mode}`);
    equal(moment(file.apptStartWhole).format('YYYY-MM-DD HH:mm'), '2026-09-24 10:00', `Exception start via ${mode}`);
    equal(moment(file.apptEndWhole).format('YYYY-MM-DD HH:mm'), '2026-09-24 11:30', `Exception end via ${mode}`);
  }
  const movedLast = { ...withException, exceptionInfo: [{ ...moved, startDateTime: minutes('2026-09-30') + 540, endDateTime: minutes('2026-09-30') + 720 }] };
  equal(plugin.findClosestOccurrence(movedLast, moment('2025-10-31T09:00:00'), moment('2026-10-01T12:00:00'))?.format('YYYY-MM-DD HH:mm'), '2026-09-30 09:00', 'Moved final date beyond nominal end');
  const sameDay = makeMeeting();
  sameDay.apptRecur = { ...withException, exceptionInfo: [{ ...moved, startDateTime: minutes('2026-09-25') + 900, endDateTime: minutes('2026-09-25') + 945 }] };
  equal(await plugin.correctRecurringOccurrenceDate(sameDay), true, 'Same-day time exception');
  equal(moment(sameDay.apptStartWhole).format('YYYY-MM-DD HH:mm'), '2026-09-25 15:00', 'Changed start time');
  equal(moment(sameDay.apptEndWhole).format('YYYY-MM-DD HH:mm'), '2026-09-25 15:45', 'Changed duration');
  const yearly = { ...active, recurFrequency: 8205, patternType: PatternType.Month, period: 12, patternTypeMonth: { day: 29 }, startDate: minutes('2025-02-28') };
  equal(find(yearly, '2025-02-28T09:00:00', '2028-03-01T12:00:00'), '2028-02-29 09:00', 'Restore leap day after non-leap start');
  const monthEnd = { ...active, patternType: PatternType.MonthEnd, startDate: minutes('2026-04-30') };
  equal(find(monthEnd, '2026-04-30T09:00:00', '2026-06-01T12:00:00'), '2026-05-31 09:00', 'MonthEnd without redundant day field');
  equal(find({ ...pattern, patternTypeMonthNth: { dayOfWeekBits: 0, n: 5 } }, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), undefined, 'Empty weekday mask');
  equal(find({ ...pattern, patternTypeMonthNth: { dayOfWeekBits: 32, n: 0 } }, '2025-10-31T09:00:00', '2026-10-01T12:00:00'), undefined, 'Invalid ordinal');
  const loneMoved = { recurrencePattern: { ...pattern, endDate: minutes('2025-10-31'), deletedInstanceDates: [minutes('2025-10-31')] }, exceptionInfo: [{ originalStartTime: minutes('2025-10-31') + 540, startDateTime: minutes('2025-10-30') + 600, endDateTime: minutes('2025-10-30') + 690, overrideFlags: 0 }] };
  equal(plugin.findClosestOccurrence(loneMoved, moment('2025-10-31T09:00:00'), moment('2025-11-01T12:00:00'))?.format('YYYY-MM-DD HH:mm'), '2025-10-30 10:00', 'Moved first date before nominal start');
  if (process.env.TZ === 'America/Toronto') {
    for (const date of ['2026-03-08', '2026-11-01']) {
      const sunday = { ...weekly, patternTypeWeek: { dayOfWeekBits: 1 }, startDate: minutes('2026-01-04') };
      equal(find(sunday, '2026-01-04T23:30:00', `${date}T23:45:00`), `${date} 23:30`, 'Late Sunday across DST must stay on Sunday');
    }
  }
  process.stdout.write(`${checks} edge-case assertions passed (${process.env.TZ}).\n`);

}
main().catch(error => { process.stderr.write(`${error.stack}\n`); process.exitCode = 1; });
