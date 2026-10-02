'use strict';

// Run from the repository: node test-recurrence.cjs
// Uses the real methods from main.ts; no copied recurrence algorithm or Obsidian runtime.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createRequire } = require('node:module');

function loadMethods(repo, names, Modal = class {}) {
  const localRequire = createRequire(path.join(repo, 'package.json'));
  const moment = localRequire('moment');
  const ts = localRequire('typescript');
  const { PatternType } = localRequire('@kenjiuno/msgreader');
  const source = ts.createSourceFile('main.ts', fs.readFileSync(path.join(repo, 'main.ts'), 'utf8'), ts.ScriptTarget.Latest, true);
  const plugin = source.statements.find(n => ts.isClassDeclaration(n) && n.name?.text === 'OutlookMeetingNotes');
  assert.ok(plugin, 'Plugin class must exist');
  const methods = names.map(name => {
    const method = plugin.members.find(n => n.name?.getText(source) === name);
    assert.ok(method, `Missing method: ${name}`);
    return method.getText(source);
  });
  const allMethods = plugin.members.filter(n => ts.isMethodDeclaration(n)).map(n => n.getText(source)).join('\n');
  const code = ts.transpileModule(
    `module.exports = (moment, PatternType, OccurrenceDateModal, ICAL, findIana) => new class { ${allMethods} };`,
    { compilerOptions: { target: ts.ScriptTarget.ES2021, module: ts.ModuleKind.CommonJS } }
  ).outputText;
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-recurrence-'));
  const file = path.join(temp, 'methods.cjs');
  try {
    fs.writeFileSync(file, code);
    return { plugin: require(file)(moment, PatternType, Modal, localRequire('ical.js'), localRequire('windows-iana').findIana), moment, PatternType };
  } finally {
    delete require.cache[file];
    fs.unlinkSync(file);
    fs.rmdirSync(temp);
  }
}

const methodNames = ['dateFromRecurMinutes', 'nthWeekdayOfMonth', 'findClosestOccurrence'];
module.exports = { loadMethods, methodNames };

if (require.main === module) {
  // Only this Node process changes timezone, never Windows or Obsidian.
  process.env.TZ = process.env.RECURRENCE_TEST_TZ || 'America/Toronto';
  const repo = path.resolve(process.argv[2] || __dirname);
  const { plugin, moment, PatternType } = loadMethods(repo, methodNames);
  moment.locale('en');
  const minutes = date => (Date.parse(`${date}T00:00:00Z`) + 11644473600000) / 60000;
  const recurrence = {
    recurrencePattern: {
      recurFrequency: 8204, patternType: PatternType.MonthNth, period: 1,
      patternTypeMonthNth: { dayOfWeekBits: 32, n: 5 },
      startDate: minutes('2025-10-31'), endDate: minutes('2026-09-25'),
      endType: 0x2021, firstDOW: 0, deletedInstanceDates: [], modifiedInstanceDates: []
    }
  };
  const base = moment('2025-10-31T09:00:00');
  let passed = 0;
  let failed = 0;
  function check(name, expected, now, change = {}, start = base) {
    const input = { recurrencePattern: { ...recurrence.recurrencePattern, ...change } };
    const today = moment(now);
    const before = [start.valueOf(), today.valueOf(), JSON.stringify(input)];
    const actual = plugin.findClosestOccurrence(input, start, today)?.format('YYYY-MM-DD HH:mm');
    try {
      assert.equal(actual, expected);
      assert.deepEqual([start.valueOf(), today.valueOf(), JSON.stringify(input)], before, 'Inputs must not change');
      passed++;
      process.stdout.write(`PASS ${name}: ${actual}\n`);
    } catch {
      failed++;
      process.stderr.write(`FAIL ${name}: expected ${expected}, got ${actual}\n`);
    }
  }
  check('Exact reported example', '2026-09-25 09:00', '2026-10-01T12:00:00');
  check('On the final date', '2026-09-25 09:00', '2026-09-25T12:00:00');
  check('Years after the end', '2026-09-25 09:00', '2030-10-01T12:00:00');
  check('Before the series starts', '2025-10-31 09:00', '2025-09-01T12:00:00');
  check('Still active', '2026-08-28 09:00', '2026-08-27T12:00:00');
  check('No end date', '2026-10-30 09:00', '2026-10-28T12:00:00', { endDate: 0x5AE980DF, endType: 0x2023 });
  check('Ends after 12 occurrences', '2026-09-25 09:00', '2026-10-01T12:00:00', { endType: 0x2022, occurrenceCount: 12 });
  check('Winter end date', '2026-01-30 09:00', '2026-02-02T12:00:00', { endDate: minutes('2026-01-30') });
  check('Five-Friday month', '2026-07-31 09:00', '2026-08-02T12:00:00', { endDate: minutes('2026-07-31') });
  check('Final occurrence at midnight', '2026-09-25 00:00', '2026-10-01T12:00:00', {}, base.clone().hour(0));
  check('Final occurrence late at night', '2026-09-25 23:30', '2026-10-01T12:00:00', {}, base.clone().hour(23).minute(30));
  check('Daily final occurrence', '2026-09-25 09:00', '2026-10-01T12:00:00', { recurFrequency: 8202, patternType: PatternType.Day, period: 1440 });
  check('Weekly final occurrence', '2026-09-25 09:00', '2026-10-01T12:00:00', { recurFrequency: 8203, patternType: PatternType.Week, period: 1, patternTypeWeek: { dayOfWeekBits: 32 } });
  check('Monthly fixed day', '2026-09-25 09:00', '2026-10-01T12:00:00', { patternType: PatternType.Month, patternTypeMonth: { day: 25 }, startDate: minutes('2025-10-25') }, moment('2025-10-25T09:00:00'));
  check('Yearly final occurrence', '2026-10-30 09:00', '2026-11-02T12:00:00', { recurFrequency: 8205, period: 12, endDate: minutes('2026-10-30') });
  check('Every two months', '2026-08-28 09:00', '2026-10-01T12:00:00', { period: 2, endDate: minutes('2026-08-28') });
  process.stdout.write(`${passed} passed, ${failed} failed (${process.env.TZ}; Moment ${moment.version})\n`);
  process.exitCode = failed ? 1 : 0;
}
