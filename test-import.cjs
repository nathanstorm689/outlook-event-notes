'use strict';
// Audits the actual main.ts with simulated Obsidian/file APIs. No Outlook UI is simulated as a real test.
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
process.env.TZ = process.env.AUDIT_TZ || 'America/Toronto';
const repo = path.resolve(process.argv[2] || __dirname);
const req = createRequire(path.join(repo, 'package.json'));
const ts = req('typescript');
const moment = req('moment');
const yaml = req('yaml');
const { PatternType } = req('@kenjiuno/msgreader');
const MsgReader = req('@kenjiuno/msgreader').default;
const Mustache = req('mustache');
const proxyData = req('mustache-validator').default;
const ICAL = req('ical.js');
const { findIana } = req('windows-iana');
moment.locale('en');
moment.suppressDeprecationWarnings = true;
moment.now = () => new Date(2026, 9, 1, 12).valueOf();
const notices = [];
const pending = [];
const asyncErrors = [];
let modalDate;
let modalAnswer;
class Notice { constructor(message) { notices.push(message); } }
class OccurrenceDateModal {
  constructor(app, date, resolve) { modalDate = date; this.resolve = resolve; }
  open() { this.resolve(modalAnswer === undefined ? modalDate : modalAnswer); }
}
class FileReader {
  readAsText(file) { this.read(file); }
  readAsArrayBuffer(file) { this.read(file); }
  read(file) {
    this.result = file.result;
    pending.push(Promise.resolve().then(() => file.fail ? this.onerror?.({}) : this.onload?.({})).catch(e => asyncErrors.push(e.message)));
  }
}
const source = ts.createSourceFile('main.ts', fs.readFileSync(path.join(repo, 'main.ts'), 'utf8'), ts.ScriptTarget.Latest, true);
const klass = source.statements.find(n => ts.isClassDeclaration(n) && n.name?.text === 'OutlookMeetingNotes');
const members = klass.members.filter(n => ts.isMethodDeclaration(n)).map(n => n.getText(source)).join('\n');
const constants = source.statements.filter(n => ts.isVariableStatement(n) && n.declarationList.declarations.some(d => ['OutlookMeetingNotesDefaultFilenamePattern', 'OutlookMeetingNotesDefaultTemplate', 'DEFAULT_SETTINGS'].includes(d.name.getText(source)))).map(n => n.getText(source)).join('\n');
const code = ts.transpileModule(`module.exports = ({moment, PatternType, MsgReader, Mustache, proxyData, Notice, FileReader, OccurrenceDateModal, normalizePath, ICAL, findIana}) => { ${constants}\nreturn { plugin: new class { ${members} }, settings: DEFAULT_SETTINGS }; };`, {compilerOptions:{target:ts.ScriptTarget.ES2021,module:ts.ModuleKind.CommonJS}}).outputText;
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-audit-'));
const compiled = path.join(temp, 'methods.cjs');
let loaded;
try {
  fs.writeFileSync(compiled, code);
  loaded = require(compiled)({moment, PatternType, MsgReader, Mustache, proxyData, Notice, FileReader, OccurrenceDateModal, ICAL, findIana, normalizePath: s => s.replace(/\\/g, '/').replace(/\/+/g, '/').replace(/^\//, '').replace(/\/$/, '')});
} finally { delete require.cache[compiled]; fs.unlinkSync(compiled); fs.rmdirSync(temp); }
const plugin = loaded.plugin;
const results = [];
async function test(group, name, action) {
  notices.length = 0; asyncErrors.length = 0; modalAnswer = undefined; modalDate = undefined;
  try { await action(); results.push({group, name, status:'PASS'}); }
  catch (e) { results.push({group, name, status:'FAIL', detail:e.message}); }
}
const minutes = date => (Date.parse(`${date}T00:00:00Z`) + 11644473600000) / 60000;
const pattern = {recurFrequency:8204,patternType:PatternType.MonthNth,period:1,patternTypeMonthNth:{dayOfWeekBits:32,n:5},startDate:minutes('2025-10-31'),endDate:minutes('2026-09-25'),endType:0x2021,firstDOW:0,deletedInstanceDates:[],modifiedInstanceDates:[]};
const meeting = extra => ({ dataType:'msg',messageClass:'IPM.Appointment',subject:'Codex test meeting',body:'Test only',apptLocation:'Test room',recipients:[],apptStartWhole:moment('2025-10-31T09:00:00').toISOString(),apptEndWhole:moment('2025-10-31T12:00:00').toISOString(),apptRecur:{recurrencePattern:{...pattern}},...extra });
const goid = (year, month, day) => '040000008200E00074C5B7101A82E008' + year.toString(16).padStart(4,'0') + month.toString(16).padStart(2,'0') + day.toString(16).padStart(2,'0') + '0'.repeat(40);
const event = (properties, more = '') => `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Codex local audit//EN\r\nBEGIN:VEVENT\r\nUID:codex-audit@local.invalid\r\nDTSTAMP:20261001T120000Z\r\nSUMMARY:Codex test\r\n${properties}\r\nEND:VEVENT\r\n${more}END:VCALENDAR\r\n`;
const iso = date => moment(date).toISOString();
const basic = 'DTSTART:20260925T130000Z\r\nDTEND:20260925T160000Z';
function vault(settings = {}) {
  const files = new Map(); const folders = new Set(); const opened = [];
  plugin.settings = {...loaded.settings, ...settings};
  plugin.app = {vault:{
    getFileByPath: p => files.get(p) || null, getFolderByPath: p => folders.has(p) ? {} : null,
    createFolder: async p => { await Promise.resolve(); if(folders.has(p)) throw new Error('Folder already exists'); folders.add(p); },
    create: async (p, content) => { await Promise.resolve(); if(files.has(p)) throw new Error('File already exists'); const f={path:p,basename:path.basename(p,'.md'),content}; files.set(p,f); return f; }
  },workspace:{getLeaf:()=>({openFile:async f=>opened.push(f.path)})}};
  return {files,folders,opened};
}
async function main() {
  for (const [name, y, m, d, expected] of [['Normal',2026,9,25,'2026-09-25'],['Leap day',2028,2,29,'2028-02-29'],['Master',0,0,0,null],['Invalid February',2026,2,30,null],['Invalid month',2026,13,1,null]]) {
    await test('GlobalObjectId',name,()=>assert.equal(plugin.getOccurrenceDateFromGlobalId(goid(y,m,d))?.format('YYYY-MM-DD') ?? null,expected));
  }
  await test('GlobalObjectId','Non-hex bytes rejected',()=>assert.equal(plugin.getOccurrenceDateFromGlobalId('0'.repeat(32)+'zzzzzzzz'),null));
  await test('GlobalObjectId','Short identifier rejected',()=>assert.equal(plugin.getOccurrenceDateFromGlobalId('0123'),null));
  const textCases=[
    ['English long','Start: Friday, September 25, 2026 9:00 AM','2026-09-25 09:00'],
    ['French long','Début : vendredi 25 septembre 2026 09:00','2026-09-25 09:00'],
    ['English numeric','Start: 9/25/2026 9:00 AM','2026-09-25 09:00'],
    ['French numeric padded','Début : 25/09/2026 09:00','2026-09-25 09:00'],
    ['Leading whitespace','\tStart: Friday, September 25, 2026 9:00 AM','2026-09-25 09:00'],
    ['No start line','Subject: Meeting',null],['Empty start','Start: ',null],
    ['Invalid date','Start: February 30, 2026 9:00 AM',null]
  ];
  for(const [name,text,expected] of textCases) await test('Drag text',name,()=>assert.equal(plugin.parseDateFromDragText(text)?.format('YYYY-MM-DD HH:mm') ?? null,expected));
  await test('Drag text','Locale restored',()=>{ moment.locale('fr'); plugin.parseDateFromDragText(textCases[0][1]); assert.equal(moment.locale(),'fr'); moment.locale('en'); });
  moment.locale('en');
  const icsCases=[
    ['UTC timestamps',basic,'2026-09-25T13:00:00.000Z','2026-09-25T16:00:00.000Z'],
    ['Floating local','DTSTART:20260925T090000\r\nDTEND:20260925T120000',iso('2026-09-25T09:00:00'),iso('2026-09-25T12:00:00')],
    ['All day explicit end','DTSTART;VALUE=DATE:20260925\r\nDTEND;VALUE=DATE:20260926',iso('2026-09-25'),iso('2026-09-26')],
    ['All day default end','DTSTART;VALUE=DATE:20260925',iso('2026-09-25'),iso('2026-09-26')],
    ['Duration property','DTSTART:20260925T130000Z\r\nDURATION:PT3H','2026-09-25T13:00:00.000Z','2026-09-25T16:00:00.000Z'],
    ['Zero duration by default','DTSTART:20260925T130000Z','2026-09-25T13:00:00.000Z','2026-09-25T13:00:00.000Z'],
    ['Paris timezone','DTSTART;TZID=Europe/Paris:20260925T090000\r\nDTEND;TZID=Europe/Paris:20260925T120000','2026-09-25T07:00:00.000Z','2026-09-25T10:00:00.000Z'],
    ['Windows Pacific timezone','DTSTART;TZID="Pacific Standard Time":20260925T090000\r\nDTEND;TZID="Pacific Standard Time":20260925T120000','2026-09-25T16:00:00.000Z','2026-09-25T19:00:00.000Z']
  ];
  for(const [name,props,start,end] of icsCases) await test('ICS dates',name,()=>{const f=plugin.parseIcsFile(event(props)); assert.equal(f.apptStartWhole,start); assert.equal(f.apptEndWhole,end);});
  for(const [name,props] of [['Invalid date','DTSTART:20260230T090000'],['Garbage suffix','DTSTART:20260925T090000GARBAGE'],['End before start','DTSTART:20260925T130000Z\r\nDTEND:20260925T120000Z'],['Missing start','DTEND:20260925T130000Z']])
    await test('ICS validation',name,()=>assert.throws(()=>plugin.parseIcsFile(event(props))));
  await test('ICS validation','No VEVENT',()=>assert.throws(()=>plugin.parseIcsFile('BEGIN:VCALENDAR\nEND:VCALENDAR')));
  await test('ICS content','Folded Unicode title',()=>{const f=plugin.parseIcsFile(event(basic).replace('SUMMARY:Codex test','SUMMARY:Réunion de\r\n suivi 😊')); assert.equal(f.subject,'Réunion desuivi 😊');});
  await test('ICS content','Text newline and punctuation escapes',()=>assert.equal(plugin.parseIcsFile(event(basic+'\r\nDESCRIPTION:Line one\\nLine two\\, item\\; end')).body,'Line one\nLine two, item; end'));
  await test('ICS content','Literal backslash-n',()=>assert.equal(plugin.parseIcsFile(event(basic+'\r\nDESCRIPTION:literal \\\\n text')).body,'literal \\n text'));
  for(const cn of ['Doe; Jane','Doe: Jane']) await test('ICS content',`Quoted attendee ${cn}`,()=>assert.deepEqual(plugin.parseIcsFile(event(basic+`\r\nATTENDEE;CN="${cn}":mailto:test@example.invalid`)).recipients,[{name:cn,email:'test@example.invalid'}]));
  await test('ICS content','Alarm description is not event body',()=>assert.equal(plugin.parseIcsFile(event(basic+'\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder only\r\nTRIGGER:-PT15M\r\nEND:VALARM')).body,''));
  await test('ICS recurrence','Master RRULE must not silently become a single event',()=>{const f=plugin.parseIcsFile(event('DTSTART:20251031T130000Z\r\nDTEND:20251031T160000Z\r\nRRULE:FREQ=MONTHLY;BYDAY=-1FR;UNTIL=20260925T130000Z')); assert.ok(f.apptRecur || f.needsOccurrenceDate,'Recurring series has no occurrence-selection signal');});
  await test('ICS recurrence','Explicit RECURRENCE-ID occurrence',()=>assert.equal(plugin.parseIcsFile(event(basic+'\r\nRECURRENCE-ID:20260925T130000Z')).apptStartWhole,'2026-09-25T13:00:00.000Z'));
  await test('ICS recurrence','Multiple events must not be silently discarded',()=>assert.throws(()=>plugin.parseIcsFile(event(basic,'BEGIN:VEVENT\r\nUID:other@local.invalid\r\nSUMMARY:Other meeting\r\nDTSTART:20260926T130000Z\r\nEND:VEVENT\r\n'))));
  await test('Recurrence','Every weekday encoded as Daily + Week',()=>{const rp={...pattern,recurFrequency:8202,patternType:PatternType.Week,patternTypeWeek:{dayOfWeekBits:62},period:1,endDate:0x5AE980DF}; assert.equal(plugin.findClosestOccurrence({recurrencePattern:rp},moment('2025-10-31T09:00:00'),moment('2026-09-26T12:00:00'))?.format('YYYY-MM-DD HH:mm'),'2026-09-25 09:00');});
  await test('Recurrence','Invalid global ID falls through to modal',async()=>{const f=meeting({globalAppointmentID:goid(2026,2,30)}); await plugin.correctRecurringOccurrenceDate(f); assert.equal(f.apptStartWhole,iso('2026-09-25T09:00:00'));});
  await test('Recurrence','All-day event spanning spring DST keeps one calendar day',async()=>{const f=meeting({apptStartWhole:iso('2026-03-01'),apptEndWhole:iso('2026-03-02'),globalAppointmentID:goid(2026,3,8),apptSubType:1,apptRecur:{recurrencePattern:{...pattern,recurFrequency:8203,patternType:PatternType.Week,period:1,startDate:minutes('2026-03-01'),endDate:0x5AE980DF,patternTypeWeek:{dayOfWeekBits:1}}}}); await plugin.correctRecurringOccurrenceDate(f); assert.equal(f.apptEndWhole,iso('2026-03-09'));});
  await test('Recurrence timezone','Paris event follows Paris DST rules',async()=>{const f=meeting({apptStartWhole:'2025-10-31T08:00:00.000Z',apptEndWhole:'2025-10-31T11:00:00.000Z',timeZoneDesc:'Romance Standard Time',globalAppointmentID:goid(2026,9,25),timeZoneStruct:{bias:-60,standardBias:0,daylightBias:-60,standardYear:0,daylightYear:0,standardDate:{year:0,month:10,dayOfWeek:0,day:5,hour:3,minute:0},daylightDate:{year:0,month:3,dayOfWeek:0,day:5,hour:2,minute:0}}}); await plugin.correctRecurringOccurrenceDate(f); assert.equal(f.apptStartWhole,'2026-09-25T07:00:00.000Z');});
  await test('Recurrence timezone','Series master on previous local date must not bypass correction',async()=>{const f=meeting({apptStartWhole:'2025-10-30T11:00:00.000Z',apptEndWhole:'2025-10-30T12:00:00.000Z',timeZoneDesc:'New Zealand Standard Time',globalAppointmentID:goid(2026,9,25)}); await plugin.correctRecurringOccurrenceDate(f); assert.equal(f.apptStartWhole,'2026-09-24T12:00:00.000Z');});
  for(const [name,subject] of [['Quoted colon','Topic: "quoted" \\path'],['Boolean title','true'],['Null title','null'],['Numeric title','12345'],['Unicode title','Réunion 😊']])
    await test('Note/YAML',name,()=>{const f=meeting({subject,apptRecur:null}); plugin.addHelperFunctions(f); const text=plugin.renderTemplate(loaded.settings.notesTemplate,f); const front=text.split('---')[1]; assert.equal(yaml.parse(front).title,subject);});
  await test('Note/YAML','Recipient name stays text',()=>{const f=meeting({recipients:[{name:'null',email:'test@example.invalid'}]}); plugin.addHelperFunctions(f); assert.deepEqual(yaml.parse(plugin.renderTemplate(loaded.settings.notesTemplate,f).split('---')[1])['meeting-recipients'],['null']);});
  await test('Vault simulation','Second import opens existing note',async()=>{const v=vault(); await plugin.createNoteFromFileData(meeting({apptRecur:null})); await plugin.createNoteFromFileData(meeting({apptRecur:null})); assert.equal(v.files.size,1); assert.equal(v.opened.length,2);});
  await test('Vault simulation','Simultaneous identical imports',async()=>{const v=vault(); const r=await Promise.allSettled([plugin.createNoteFromFileData(meeting({apptRecur:null})),plugin.createNoteFromFileData(meeting({apptRecur:null}))]); assert.ok(r.every(x=>x.status==='fulfilled'),JSON.stringify(r.map(x=>x.reason?.message))); assert.equal(v.files.size,1);});
  await test('Vault simulation','Reserved Windows filename',async()=>{const v=vault({fileNamePattern:'{{subject}}'}); await plugin.createNoteFromFileData(meeting({apptRecur:null,subject:'CON'})); assert.ok(![...v.files.keys()].some(p=>/^CON\.md$/i.test(p)),'Generated CON.md');});
  await test('Vault simulation','Filename control characters',async()=>{const v=vault(); await plugin.createNoteFromFileData(meeting({apptRecur:null,subject:'Line one\nLine two'})); assert.ok(![...v.files.keys()].some(p=>/[\x00-\x1f]/.test(p)),'Filename contains a newline');});
  await test('MSG validation','Mail item rejected',async()=>{vault(); await assert.rejects(()=>plugin.createMeetingNote({getFileData:()=>meeting({messageClass:'IPM.Note'})}));});
  await test('MSG validation','Meeting invitation message rejected clearly',async()=>{vault(); await assert.rejects(()=>plugin.createMeetingNote({getFileData:()=>meeting({messageClass:'IPM.Schedule.Meeting.Request'})})); assert.ok(notices.length);});
  await test('MSG validation','Missing start must not create an undated note',async()=>{const v=vault(); try {await plugin.createMeetingNote({getFileData:()=>meeting({apptStartWhole:undefined,apptRecur:null})});} catch {} assert.equal(v.files.size,0);});
  await test('Drop simulation','No file shows notice',()=>{plugin.handleDropEvent({dataTransfer:{files:[]}}); assert.equal(notices.length,1);});
  await test('Drop simulation','Multiple files show notice',()=>{plugin.handleDropEvent({dataTransfer:{files:[{},{}]}}); assert.equal(notices.length,1);});
  await test('Drop simulation','Uppercase ICS recognized',async()=>{const v=vault(); plugin.handleDropEvent({dataTransfer:{files:[{name:'TEST.ICS',type:'',result:event(basic)}]}}); await Promise.all(pending.splice(0)); assert.equal(v.files.size,1);});
  await test('Drop simulation','Unreadable file shows notice',async()=>{vault(); plugin.handleDropEvent({dataTransfer:{files:[{name:'test.msg',type:'',fail:true}],getData:()=>''}}); await Promise.all(pending.splice(0)); assert.ok(notices.length,'No FileReader error notice');});
  await test('Drop simulation','Corrupt MSG is caught at async boundary',async()=>{vault(); plugin.handleDropEvent({dataTransfer:{files:[{name:'broken.msg',type:'',result:new ArrayBuffer(4)}],getData:()=>''}}); await Promise.all(pending.splice(0)); assert.equal(asyncErrors.length,0,asyncErrors.join('; ')); assert.ok(notices.length);});
  await test('Regression/YAML','Double-quoted custom template',()=>{const f=meeting({subject:'Hello: "World" \\path',apptRecur:null}); assert.equal(yaml.parse(plugin.renderTemplate('---\ntitle: "{{subject}}"\n---\n',f).split('---')[1]).title,f.subject);});
  await test('Regression/YAML','Single-quoted custom template',()=>{const f=meeting({subject:"O'Brien: meeting",apptRecur:null}); assert.equal(yaml.parse(plugin.renderTemplate("---\ntitle: '{{subject}}'\n---\n",f).split('---')[1]).title,f.subject);});
  await test('Regression/YAML','Multiline value remains text',()=>{const f=meeting({body:'one\ntwo: "quoted"'}); assert.equal(yaml.parse(plugin.renderTemplate('---\nbody: {{body}}\n---\n',f).split('---')[1]).body,f.body);});
  await test('Regression/YAML','First-word helper still works',()=>{const f=meeting({subject:'Hello World'}); plugin.addHelperFunctions(f); assert.equal(yaml.parse(plugin.renderTemplate('---\nword: {{#helper_firstWord}}{{subject}}{{/helper_firstWord}}\n---\n',f).split('---')[1]).word,'Hello');});
  await test('Regression/YAML','Helper result stays text',()=>{const f=meeting({subject:'true value'}); plugin.addHelperFunctions(f); assert.equal(yaml.parse(plugin.renderTemplate('---\nword: {{#helper_firstWord}}{{subject}}{{/helper_firstWord}}\n---\n',f).split('---')[1]).word,'true');});
  await test('Regression/YAML','Quoted key does not change helper escaping',()=>{const f=meeting({subject:'true value'}); plugin.addHelperFunctions(f); assert.equal(yaml.parse(plugin.renderTemplate("---\n'word': {{#helper_firstWord}}{{subject}}{{/helper_firstWord}}\n---\n",f).split('---')[1]).word,'true');});
  const invitationBodies = [
    ['Leading padded blank line (CRLF)', '    \r\nBonjour,\r\n\r\nPour la visite.\r\n'],
    ['Leading tab-only line (LF)', '\t \nBonjour,\n\nPour la visite.\n'],
    ['Leading padded blank line (CR)', '    \rBonjour,\r\rPour la visite.\r'],
    ['Mixed blank-line indentation', ' \r\n       \r\n Bonjour,\r\n\r\nPour la visite.\r\n'],
    ['Invitation containing a separator', '    \r\nBonjour,\r\n---\r\nPour la visite.\r\n']
  ];
  for (const [name, body] of invitationBodies) await test('Regression/Invitation whitespace', name, async () => {
    const v = vault();
    await plugin.createNoteFromFileData(meeting({ apptRecur: null, body }));
    assert.equal(v.files.size, 1);
    const note = [...v.files.values()][0].content;
    const front = note.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
    assert.ok(front, 'Complete frontmatter');
    assert.equal(yaml.parse(front[1])['meeting-invite'], body, 'Invitation whitespace and paragraphs preserved');
  });
  await test('Regression/ICS','Unknown timezone rejected',()=>assert.throws(()=>plugin.parseIcsFile(event('DTSTART;TZID=Not/AZone:20260925T090000'))));
  await test('Regression/ICS','Embedded custom timezone',()=>{const ics=event('DTSTART;TZID=Custom:20260925T090000\r\nDTEND;TZID=Custom:20260925T120000').replace('BEGIN:VEVENT','BEGIN:VTIMEZONE\r\nTZID:Custom\r\nBEGIN:STANDARD\r\nDTSTART:19700101T000000\r\nTZOFFSETFROM:+0230\r\nTZOFFSETTO:+0230\r\nEND:STANDARD\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT'); const f=plugin.parseIcsFile(ics); assert.equal(f.apptStartWhole,'2026-09-25T06:30:00.000Z'); assert.equal(f.apptEndWhole,'2026-09-25T09:30:00.000Z');});
  await test('Regression/ICS','Nominal day duration crosses DST',()=>{const f=plugin.parseIcsFile(event('DTSTART;TZID=America/Toronto:20260307T120000\r\nDURATION:P1D')); assert.equal(f.apptEndWhole,'2026-03-08T16:00:00.000Z');});
  await test('Regression/ICS','Exact hour duration crosses DST',()=>{const f=plugin.parseIcsFile(event('DTSTART;TZID=America/Toronto:20260308T010000\r\nDURATION:PT3H')); assert.equal(f.apptEndWhole,'2026-03-08T09:00:00.000Z');});
  await test('Regression/ICS','Spring missing time uses earlier offset',()=>assert.equal(plugin.parseIcsFile(event('DTSTART;TZID=America/Toronto:20260308T023000')).apptStartWhole,'2026-03-08T07:30:00.000Z'));
  await test('Regression/ICS','Autumn repeated time uses first instance',()=>assert.equal(plugin.parseIcsFile(event('DTSTART;TZID=America/Toronto:20261101T013000')).apptStartWhole,'2026-11-01T05:30:00.000Z'));
  await test('Regression/ICS','Recurring master confirms selected local date',async()=>{const v=vault(); modalAnswer='2026-09-25'; const f=plugin.parseIcsFile(event('DTSTART:20251031T130000Z\r\nDTEND:20251031T160000Z\r\nRRULE:FREQ=MONTHLY;BYDAY=-1FR')); await plugin.createNoteFromFileData(f); assert.equal(modalDate,''); assert.equal(v.files.size,1); const expected=['2026-09-24T13:00:00.000Z','2026-09-25T13:00:00.000Z','2026-09-26T13:00:00.000Z'].find(s=>{const d=new Date(s);return d.getFullYear()===2026&&d.getMonth()===8&&d.getDate()===25;}); assert.equal(f.apptStartWhole,expected);});
  await test('Regression/ICS','Cancel series import writes nothing',async()=>{const v=vault(); modalAnswer=null; await plugin.createNoteFromFileData(plugin.parseIcsFile(event(basic+'\r\nRRULE:FREQ=DAILY'))); assert.equal(v.files.size,0);});
  await test('Regression/Vault','Concurrent imports creating a folder',async()=>{const v=vault({notesFolder:'QA'}); const both=await Promise.allSettled([plugin.createNoteFromFileData(meeting({apptRecur:null,subject:'A'})),plugin.createNoteFromFileData(meeting({apptRecur:null,subject:'B'}))]); assert.ok(both.every(x=>x.status==='fulfilled')); assert.equal(v.files.size,2);});
  await test('Regression/Vault','Replacement cannot introduce forbidden filename characters',async()=>{const v=vault({invalidFilenameCharReplacement:'/:*?'}); await plugin.createNoteFromFileData(meeting({apptRecur:null,subject:'A:B'})); assert.ok(![...v.files.keys()].some(p=>/[:*?]/.test(p)));});
  const report={timestamp:new Date().toISOString(),timezone:process.env.TZ,source:'actual main.ts; synthetic fixtures and mocked Obsidian/FileReader APIs',outlookUiTested:false,summary:{passed:results.filter(r=>r.status==='PASS').length,failed:results.filter(r=>r.status==='FAIL').length},results};
  if (process.env.IMPORT_TEST_REPORT) fs.writeFileSync(process.env.IMPORT_TEST_REPORT,JSON.stringify(report,null,2));
  for(const r of results) process.stdout.write(`${r.status} [${r.group}] ${r.name}${r.detail ? ': '+r.detail.split('\n')[0] : ''}\n`);
  process.stdout.write(JSON.stringify(report.summary)+'\n');
  process.exitCode=report.summary.failed ? 1 : 0;
}
main().catch(e=>{process.stderr.write(e.stack+'\n');process.exitCode=1;});
