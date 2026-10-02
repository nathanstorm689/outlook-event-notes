import { App, displayTooltip, Modal, moment as _moment, Notice, Plugin, PluginSettingTab, Setting, SettingDefinitionItem, TooltipPlacement, normalizePath } from 'obsidian';
import MsgReader, { AppointmentRecur, FieldsData, PatternType } from '@kenjiuno/msgreader';
import proxyData from 'mustache-validator';
import Mustache from 'mustache';
import ICAL from 'ical.js';
import { findIana } from 'windows-iana';

// moment is provided by Obsidian's bundle; importing from 'obsidian' satisfies the plugin review requirement.
// The namespace-style re-export in obsidian.d.ts loses the callable signature, so we cast it here.
const moment = _moment as unknown as typeof import('moment/moment');

// MeetingFileData extends FieldsData with additional dynamic fields added at runtime.
// recipients is widened to include plain {name, email} objects from .ics parsing.
// apptRecur is widened to allow null (used by .ics path to skip occurrence correction).
type IcalComponent = InstanceType<typeof ICAL.Component>;
type IcalProperty = InstanceType<typeof ICAL.Property>;
type IcalTime = InstanceType<typeof ICAL.Time>;
type OutlookZoneTransition = NonNullable<FieldsData['timeZoneStruct']>['standardDate'];
interface MeetingClock {
	toWall(instant: moment.Moment): moment.Moment;
	fromWall(wall: moment.Moment): moment.Moment;
}

interface MeetingFileData extends Omit<FieldsData, 'recipients' | 'apptRecur'> {
	needsOccurrenceDate?: boolean;
	icsOccurrence?: { clock: MeetingClock; startWall: moment.Moment; calendarDays: number; seconds: number };
	bodyText?: string;
	bodyPlainText?: string;
	helper_currentDT?: string;
	recipients?: Array<{ name?: string; email?: string }>;
	apptRecur?: AppointmentRecur | null;
	[key: string]: unknown;
}

const OutlookMeetingNotesDefaultFilenamePattern =
	'{{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD_HH-mm-ss{{/helper_dateFormat}} {{subject}}';

const OutlookMeetingNotesDefaultTemplate = `---
title: {{subject}}
subtitle: meeting notes
created: {{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD_HH-mm-ss{{/helper_dateFormat}}
meeting: 'true'
meeting-location: {{apptLocation}}
meeting-recipients:
{{#recipients}}
  - {{name}}
{{/recipients}}
meeting-invite: {{body}}
---
`;

interface OutlookMeetingNotesSettings {
	notesFolder: string;
	invalidFilenameCharReplacement: string;
	fileNamePattern: string;
	notesTemplate: string;
}

const DEFAULT_SETTINGS: OutlookMeetingNotesSettings = {
	notesFolder: '',
	invalidFilenameCharReplacement: '',
	fileNamePattern: OutlookMeetingNotesDefaultFilenamePattern,
	notesTemplate: OutlookMeetingNotesDefaultTemplate
}

export default class OutlookMeetingNotes extends Plugin {
	settings: OutlookMeetingNotesSettings;

	// dragText: the plain-text representation of the appointment from the DataTransfer.
	// Outlook Classic puts it there when dragging from the calendar; it contains the
	// occurrence's actual "Start:" date, which lets us skip the manual-date dialog.
	async createMeetingNote(msg: MsgReader, dragText = '', notifyErrors = true) {
		try {
			const origFileData = msg.getFileData();
			if (origFileData.dataType != 'msg') {
				throw new TypeError('Outlook Event Notes cannot process the file. '
					+ 'MsgReader did not parse the file as valid msg format.');
			} else if (origFileData.messageClass != 'IPM.Appointment') {
				throw new TypeError('Outlook Event Notes cannot process the file. '
					+ 'It is a valid msg file but not an appointment or meeting.');
			}
			await this.createNoteFromFileData(origFileData as MeetingFileData, dragText);
		} catch (ee: unknown) {
			if (notifyErrors && ee instanceof Error) { new Notice('Error (' + ee.name + '):\n' + ee.message); }
			throw ee;
		}
	}

	// Shared note-creation logic used by both the .msg path and the .ics path.
	// fileData must have: subject, apptStartWhole, apptEndWhole, apptLocation,
	// body/bodyText/bodyHtml, recipients[], apptRecur (null = skip date correction).
	private async createNoteFromFileData(origFileData: MeetingFileData, dragText = ''): Promise<void> {
		const { vault } = this.app;
		const fileData = origFileData;
		this.ensureBodyField(fileData);
		this.ensureDefaultFields(fileData);
		const start = this.parseMeetingDate(fileData.apptStartWhole);
		const end = this.parseMeetingDate(fileData.apptEndWhole);
		if (!start.isValid()) throw new TypeError('The meeting has no valid start date.');
		if (!end.isValid() || end.isBefore(start)) throw new TypeError('The meeting has no valid end date.');
		if (fileData.needsOccurrenceDate && !(await this.correctIcsOccurrenceDate(fileData))) return;
		if (!(await this.correctRecurringOccurrenceDate(fileData, dragText))) return;
		this.addHelperFunctions(fileData);
		fileData.helper_currentDT = moment().format();

		let targetFolderPath = (this.settings.notesFolder ?? '').trim();
		if (targetFolderPath === '' || targetFolderPath === '/') targetFolderPath = '';
		else targetFolderPath = normalizePath(targetFolderPath);
		if (targetFolderPath.split('/').some(p => p === '..' || p === '.')) throw new TypeError('The notes folder must be inside the vault.');
		const renderedName = Mustache.render(this.settings.fileNamePattern, proxyData(fileData), undefined, { escape: (s: string) => s });
		const filePath = normalizePath((targetFolderPath ? targetFolderPath + '/' : '') + this.sanitizeNoteName(renderedName) + '.md');
		let meetingNoteFile = vault.getFileByPath(filePath);
		let created = false;
		if (!meetingNoteFile) {
			if (targetFolderPath && !vault.getFolderByPath(targetFolderPath)) {
				try { await vault.createFolder(targetFolderPath); }
				catch (error) { if (!vault.getFolderByPath(targetFolderPath)) throw error; }
			}
			try {
				meetingNoteFile = await vault.create(filePath, this.renderTemplate(this.settings.notesTemplate, fileData));
				created = true;
			} catch (error) {
				// Another drop may have completed while this one was waiting for disk I/O.
				meetingNoteFile = vault.getFileByPath(filePath);
				if (!meetingNoteFile) throw error;
			}
		}
		new Notice(created ? 'New file created: ' + meetingNoteFile.basename : meetingNoteFile.basename + ' already exists: opening it');
		await this.app.workspace.getLeaf(false).openFile(meetingNoteFile);
	}

	private ensureBodyField(fileData: MeetingFileData): void {
		const hasBodyString = typeof fileData.body === 'string' && fileData.body.trim() !== '';
		if (hasBodyString) { return; }
		const fallbacks = [
			fileData.bodyText,
			fileData.bodyPlainText,
			fileData.bodyHtml,
			fileData.rtfCompressed
		];
		for (const candidate of fallbacks) {
			if (typeof candidate === 'string' && candidate.trim() !== '') {
				fileData.body = candidate.includes('<') ? this.dropHtmlTags(candidate) : candidate;
				return;
			}
		}
		fileData.body = '';
	}

	private ensureDefaultFields(fileData: MeetingFileData): void {
		const stringDefaults = ['subject', 'apptLocation', 'apptStartWhole', 'apptEndWhole'];
		for (const field of stringDefaults) {
			if (fileData[field] == null) { fileData[field] = ''; }
		}
		if (!Array.isArray(fileData.recipients)) { fileData.recipients = []; }
	}

	private dropHtmlTags(input: string): string {
		return input
			.replace(/<(style|script)[^>]*?>[\s\S]*?<\/\1>/gi, '')
			.replace(/<[^>]+>/g, '')
			.replace(/&nbsp;/gi, ' ')
			.replace(/&amp;/gi, '&')
			.replace(/&lt;/gi, '<')
			.replace(/&gt;/gi, '>')
			.replace(/&quot;/gi, '"')
			.replace(/&#39;/gi, "'")
			.replace(/\s+/g, ' ')
			.trim();
	}

	// Handle a file being dropped onto the ribbon icon.
	// Accepts both Outlook .msg files and iCalendar .ics files.
	// Recurring series exports can require explicit occurrence selection.
	handleDropEvent(dropevt: DragEvent): void {
		const transfer = dropevt.dataTransfer;
		if (!transfer || transfer.files.length === 0) {
			new Notice('No file received. Use Outlook Classic, or export one meeting as a .msg or .ics file.');
			return;
		}
		if (transfer.files.length !== 1) { new Notice('Only one meeting file can be dropped at a time.'); return; }
		const file = transfer.files[0];
		const isIcs = /\.ics$/i.test(file.name) || file.type.toLowerCase() === 'text/calendar';
		if (!isIcs && !/\.msg$/i.test(file.name)) { new Notice('Choose a .msg or .ics meeting file.'); return; }
		const fail = (error: unknown): void => { new Notice('Unable to import meeting: ' + (error instanceof Error ? error.message : String(error))); };
		try {
			const dragText = transfer.getData?.('text/plain') ?? '';
			const reader = new FileReader();
			reader.onerror = () => fail(reader.error ?? new Error('The file could not be read.'));
			reader.onabort = () => fail(new Error('File reading was cancelled.'));
			reader.onload = async () => {
				try {
					if (isIcs) {
						if (typeof reader.result !== 'string') throw new TypeError('The calendar file could not be read as text.');
						await this.createNoteFromFileData(this.parseIcsFile(reader.result));
					} else {
						if (!(reader.result instanceof ArrayBuffer)) throw new TypeError('The Outlook file could not be read.');
						await this.createMeetingNote(new MsgReader(reader.result), dragText, false);
					}
				} catch (error) { fail(error); }
			};
			if (isIcs) reader.readAsText(file, 'utf-8');
			else reader.readAsArrayBuffer(file);
		} catch (error) { fail(error); }
	}

	private ribbonIconEl: HTMLElement;

	async onload() {
		await this.loadSettings();

		const tooltipMessage = 'Outlook Event Notes: Drag and drop a meeting onto this icon from Outlook (or a .msg file) to create a meeting note.';

		this.ribbonIconEl = this.addRibbonIcon('calendar-clock', tooltipMessage, () => { });

		this.registerDomEvent(this.ribbonIconEl, 'dragenter', () => {
			this.ribbonIconEl.toggleClass('is-being-dragged-over', true);
			const ttPosition = this.ribbonIconEl.getAttribute('data-tooltip-position') as TooltipPlacement;
			const ttDelay = this.ribbonIconEl.getAttribute('data-tooltip-delay');
			if (ttPosition != null && ttDelay != null) {
				displayTooltip(this.ribbonIconEl, tooltipMessage, { placement: ttPosition, delay: Number(ttDelay) });
			} else {
				displayTooltip(this.ribbonIconEl, tooltipMessage);
			}
		});
		this.registerDomEvent(this.ribbonIconEl, 'dragleave', () => {
			this.ribbonIconEl.toggleClass('is-being-dragged-over', false);
		});
		this.registerDomEvent(this.ribbonIconEl, 'dragover', (dragevt: DragEvent) => {
			dragevt.preventDefault();
			if (dragevt.dataTransfer != null) { dragevt.dataTransfer.dropEffect = 'copy'; }
		});
		this.registerDomEvent(this.ribbonIconEl, 'drop', (dropevt: DragEvent) => {
			this.ribbonIconEl.toggleClass('is-being-dragged-over', false);
			dropevt.preventDefault();
			this.handleDropEvent(dropevt);
		});
		this.ribbonIconEl.addClass('outlook-event-notes-icon');

		this.addSettingTab(new OutlookMeetingNotesSettingTab(this.app, this));
	}

	onunload() { }

	async loadSettings() {
		const data = (await this.loadData()) as Partial<OutlookMeetingNotesSettings> | null;
		this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	addHelperFunctions(hash: MeetingFileData): MeetingFileData {
		const helperFunctions = {
			firstWord: () => {
				return function (words: string, render: (text: string) => string) {
					const rendered = render(words);
					let raw = rendered;
					try {
						const decoded: unknown = JSON.parse(rendered);
						if (typeof decoded === 'string') raw = decoded;
					} catch { /* Plain text from the Markdown renderer. */ }
					return raw.replace(/\W.*$/, '');
				}
			},
			dateFormat: () => {
				return function (datetime_format: string, render: (text: string) => string) {
					const parts = datetime_format.split('|');
					const rawValue = render(parts[0]).trim().replace(/^"|"$/g, '');
					const formattedMoment = moment(rawValue);
					if (!formattedMoment.isValid()) { return rawValue; }
					return formattedMoment.format(parts[1]);
				}
			}
		};
		let func: 'firstWord' | 'dateFormat';
		for (func in helperFunctions) {
			hash['helper_' + func] = helperFunctions[func];
		}
		// Add helper functions to array items so they work inside mustache sections
		for (const property in hash) {
			if (hash[property] instanceof Array) {
				for (const subItem of hash[property] as MeetingFileData[]) {
					if (subItem instanceof Object) {
						for (func in helperFunctions) {
							subItem['helper_' + func] = helperFunctions[func];
						}
					}
				}
			}
		}

		return hash;
	}

	// Parse template into YAML and markdown sections to use different escaping for each
	renderTemplate(template: string, hash: MeetingFileData): string {
		const match = template.match(/^---(\r\n?|\n).*?(\r\n?|\n)---($|\r\n?|\n)/s);
		let output = '';
		if (match) {
			// A quoted Mustache field needs escaping inside its existing YAML quotes.
			const withHelpers = match[0].replace(/\{\{#(helper_firstWord|helper_dateFormat)\}\}[\s\S]*?\{\{\/\1\}\}/g, (section: string, _name: string, offset: number) => {
				const prefix = match[0].slice(match[0].lastIndexOf('\n', offset - 1) + 1, offset);
				// Leave explicitly quoted sections inside their existing YAML scalar.
				let quote = '';
				for (let i = 0; i < prefix.length; i++) {
					if (quote === '"' && prefix[i] === '\\') { i++; continue; }
					if (prefix[i] === quote) { if (quote === "'" && prefix[i + 1] === "'") i++; else quote = ''; }
					else if (!quote && (prefix[i] === '"' || prefix[i] === "'")) quote = prefix[i];
				}
				if (quote) return section;
				return '{{#helper_yamlScalar}}' + section + '{{/helper_yamlScalar}}';
			});
			const yamlTemplate = withHelpers.replace(/\{\{\s*([^#^/!>&{=][^}]*?)\s*\}\}/g, (tag: string, name: string, offset: number) => {
				const line = withHelpers.slice(withHelpers.lastIndexOf('\n', offset - 1) + 1, offset);
				let quote = '';
				for (let i = 0; i < line.length; i++) {
					if (quote === '"' && line[i] === '\\') { i++; continue; }
					if (line[i] === quote) { if (quote === "'" && line[i + 1] === "'") i++; else quote = ''; }
					else if (!quote && (line[i] === '"' || line[i] === "'")) quote = line[i];
				}
				if (!quote) return tag;
				const helper = quote === '"' ? 'helper_yamlDouble' : 'helper_yamlSingle';
				return `{{#${helper}}}{{{${name.trim()}}}}{{/${helper}}}`;
			});
			const data = { ...hash,
				helper_yamlScalar: () => (s: string, render: (s: string) => string) => JSON.stringify(render(s)),
				helper_yamlDouble: () => (s: string, render: (s: string) => string) => JSON.stringify(render(s)).slice(1, -1),
				helper_yamlSingle: () => (s: string, render: (s: string) => string) => render(s).replace(/'/g, "''").replace(/\r?\n/g, ' ')
			};
			output = Mustache.render(yamlTemplate, proxyData(data), undefined, { escape: (s: string) => JSON.stringify(String(s)) });
		}
		const markdown = match ? template.slice(match[0].length) : template;
		return output + Mustache.render(markdown, proxyData(hash), undefined, {
			escape: (s: string) => s.replace(/[\\`*_[\]{}<>()#!|^]/g, '\\$&').replaceAll('%%', '\\%\\%').replaceAll('~~', '\\~\\~').replaceAll('==', '\\=\\=')
		});
	}


	// Parse an iCalendar (.ics) file and return a fileData object compatible with
	// createNoteFromFileData. Series masters require explicit occurrence selection.
	private parseIcsFile(content: string): MeetingFileData {
		const text = content.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n').replace(/\n[ \t]/g, '');
		// Validate raw values before ICAL.Time can normalize an impossible date.
		for (const line of text.split('\n')) {
			if (!/^(DTSTART|DTEND|RECURRENCE-ID)(?:;|:)/i.test(line)) continue;
			let quoted = false;
			let colon = -1;
			for (let i = 0; i < line.length; i++) {
				if (line[i] === '"') quoted = !quoted;
				if (line[i] === ':' && !quoted) { colon = i; break; }
			}
			const raw = line.slice(colon + 1);
			const format = raw.includes('T') ? 'YYYYMMDDTHHmmss' : 'YYYYMMDD';
			if (colon < 0 || !/^\d{8}(?:T\d{6}Z?)?$/.test(raw)
				|| !moment.utc(raw.replace(/Z$/, ''), format, 'en', true).isValid()) throw new TypeError('The calendar file contains an invalid date: ' + raw);
		}
		const calendar = new ICAL.Component(ICAL.parse(text));
		if (calendar.name !== 'vcalendar') throw new TypeError('The file is not an iCalendar document.');
		const events = calendar.getAllSubcomponents('vevent');
		if (events.length !== 1) throw new TypeError('The .ics file must contain exactly one event. Export the individual meeting or occurrence.');
		const event = events[0];
		for (const name of ['dtstart', 'dtend', 'duration', 'summary', 'recurrence-id']) {
			if (event.getAllProperties(name).length > 1) throw new TypeError('The calendar file contains duplicate ' + name + ' fields.');
		}
		const startProperty = event.getFirstProperty('dtstart');
		const summary = event.getFirstPropertyValue('summary');
		if (typeof summary !== 'string') throw new TypeError('The .ics file is missing a SUMMARY (event title).');
		if (!startProperty) throw new TypeError('The .ics file is missing a DTSTART (start date).');
		const startTime = startProperty.getFirstValue();
		if (!(startTime instanceof ICAL.Time)) throw new TypeError('The calendar start date is invalid.');
		const wall = (t: IcalTime): moment.Moment => moment.utc([t.year, t.month - 1, t.day, t.hour, t.minute, t.second]);
		const clock = this.getIcsClock(calendar, startProperty, startTime);
		const startWall = wall(startTime);
		const start = clock.fromWall(startWall);
		const endProperty = event.getFirstProperty('dtend');
		const durationProperty = event.getFirstProperty('duration');
		if (endProperty && durationProperty) throw new TypeError('The calendar file must not contain both DTEND and DURATION.');
		let end: moment.Moment;
		let calendarDays = startTime.isDate ? 1 : 0;
		let seconds = 0;
		if (endProperty) {
			const endTime = endProperty.getFirstValue();
			if (!(endTime instanceof ICAL.Time) || endTime.isDate !== startTime.isDate) throw new TypeError('The calendar start and end must use the same date type.');
			end = this.getIcsClock(calendar, endProperty, endTime).fromWall(wall(endTime));
			if (startTime.isDate) calendarDays = wall(endTime).diff(startWall, 'days');
			else seconds = end.diff(start, 'seconds');
		} else if (durationProperty) {
			const duration = durationProperty.getFirstValue();
			if (!(duration instanceof ICAL.Duration) || duration.isNegative || duration.toSeconds() <= 0) throw new TypeError('The calendar duration must be positive.');
			calendarDays = duration.weeks * 7 + duration.days;
			seconds = duration.hours * 3600 + duration.minutes * 60 + duration.seconds;
			if (startTime.isDate && seconds) throw new TypeError('An all-day duration must use whole days or weeks.');
			end = clock.fromWall(startWall.clone().add(calendarDays, 'days')).add(seconds, 'seconds');
		} else end = clock.fromWall(startWall.clone().add(calendarDays, 'days'));
		if (!start.isValid() || !end.isValid() || end.isBefore(start) || (startTime.isDate && !end.isAfter(start))) throw new TypeError('The calendar end date must not precede its start date.');
		const recipients: Array<{ name: string; email: string }> = [];
		for (const attendee of event.getAllProperties('attendee')) {
			const address = attendee.getFirstValue();
			if (typeof address !== 'string' || !/^mailto:/i.test(address)) continue;
			const email = address.replace(/^mailto:/i, '');
			const name = attendee.getParameter('cn');
			recipients.push({ name: typeof name === 'string' ? name : email, email });
		}
		// DTSTART of a series master is not evidence of which occurrence was dropped.
		const recurring = !event.hasProperty('recurrence-id') && (event.hasProperty('rrule') || event.hasProperty('rdate'));
		for (const rule of event.getAllProperties('rrule')) rule.getFirstValue();
		const location = event.getFirstPropertyValue('location');
		const body = event.getFirstPropertyValue('description');
		return {
			dataType: 'msg', messageClass: 'IPM.Appointment', subject: summary,
			apptStartWhole: start.toISOString(), apptEndWhole: end.toISOString(),
			apptLocation: typeof location === 'string' ? location : '', body: typeof body === 'string' ? body : '',
			recipients, apptRecur: null, needsOccurrenceDate: recurring,
			icsOccurrence: recurring ? { clock, startWall, calendarDays, seconds } : undefined
		};
	}

	// Convert Outlook recurrence minutes (since midnight Jan 1, 1601, local time) to a moment.
	private dateFromRecurMinutes(minutes: number): moment.Moment {
		return moment(new Date(-11644473600000 + minutes * 60000));
	}

	// For recurring events, apptStartWhole stores the first occurrence's date.
	// When a later occurrence is dragged, we try to correct it using:
	//   1. PidLidGlobalObjectId bytes 16-19, which Outlook sets to the specific
	//      occurrence's year/month/day (zeros = series master / non-specific).
	//   2. If still unknown, show a date-picker dialog pre-filled with the
	//      occurrence from the recurrence pattern closest to today.
	// Returns false if the user cancelled the dialog (caller should abort note creation).
	private async correctRecurringOccurrenceDate(fileData: MeetingFileData, dragText = ''): Promise<boolean> {
		const apptRecur = fileData.apptRecur;
		if (!apptRecur?.recurrencePattern) return true;
		const rp = apptRecur.recurrencePattern;
		const apptStart = this.parseMeetingDate(fileData.apptStartWhole);
		const apptEnd = this.parseMeetingDate(fileData.apptEndWhole);
		if (!apptStart.isValid()) throw new TypeError('The meeting has no valid start date.');
		const clock = this.getMeetingClock(fileData);
		const startWall = clock.toWall(apptStart);
		const firstDate = this.dateFromRecurMinutes(rp.startDate).utc();
		if (!firstDate.isValid()) throw new TypeError('The recurring series has no valid start date.');
		// Compare calendar dates in the MEETING's zone, not the computer's zone.
		if (!startWall.isSame(firstDate, 'day')) return true;
		const withDate = (date: moment.Moment): moment.Moment => moment.utc([
			date.year(), date.month(), date.date(), startWall.hour(), startWall.minute(), startWall.second(), startWall.millisecond()
		]);
		const fromId = this.getOccurrenceDateFromGlobalId(fileData.globalAppointmentID);
		let chosen: moment.Moment;
		if (fromId) chosen = withDate(fromId);
		else {
			const parsed = dragText ? this.parseDateFromDragText(dragText) : null;
			// Outlook's text is displayed in the computer's zone. Keep its actual instant.
			if (parsed && !clock.toWall(parsed).isSame(firstDate, 'day')) chosen = clock.toWall(parsed);
			else {
				const now = moment();
				const closest = this.findClosestOccurrence(apptRecur, startWall, clock.toWall(now), d => clock.fromWall(d).diff(now));
				// The picker shows a date in the user's local calendar, just like Outlook.
				const suggestion = closest ? clock.fromWall(closest).local().locale('en').format('YYYY-MM-DD') : '';
				const selected = await new Promise<string | null>(resolve => new OccurrenceDateModal(this.app, suggestion, resolve).open());
				if (!selected) return false;
				const date = moment(selected, 'YYYY-MM-DD', 'en', true);
				if (!date.isValid()) return false;
				if (closest && selected === suggestion) chosen = closest;
				else {
					// Find the meeting-zone date which falls on the selected local date.
					const candidate = withDate(date);
					chosen = [-1, 0, 1].map(n => candidate.clone().add(n, 'days')).find(d => clock.fromWall(d).local().isSame(date, 'day')) ?? candidate;
				}
			}
		}
		const exceptions = apptRecur.exceptionInfo ?? [];
		const exception = (fromId ? exceptions.find(e => this.dateFromRecurMinutes(e.originalStartTime).utc().isSame(withDate(fromId), 'day')) : undefined)
			?? exceptions.find(e => this.dateFromRecurMinutes(e.startDateTime).utc().isSame(chosen, 'day'));
		if (exception) {
			const start = clock.fromWall(this.dateFromRecurMinutes(exception.startDateTime).utc());
			const end = clock.fromWall(this.dateFromRecurMinutes(exception.endDateTime).utc());
			if (!start.isValid() || !end.isValid() || end.isBefore(start)) throw new TypeError('The modified occurrence has invalid dates.');
			fileData.apptStartWhole = start.toISOString();
			fileData.apptEndWhole = end.toISOString();
			return true;
		}
		const endWall = apptEnd.isValid() ? clock.toWall(apptEnd) : startWall;
		const wallDuration = Number.isFinite(apptRecur.startTimeOffset) && Number.isFinite(apptRecur.endTimeOffset)
			? apptRecur.endTimeOffset - apptRecur.startTimeOffset : endWall.diff(startWall, 'minutes');
		if (wallDuration < 0) throw new TypeError('The meeting ends before it starts.');
		fileData.apptStartWhole = clock.fromWall(chosen).toISOString();
		fileData.apptEndWhole = clock.fromWall(chosen.clone().add(wallDuration, 'minutes')).toISOString();
		return true;
	}

	// Parse the occurrence date from PidLidGlobalObjectId (as hex string).
	// Bytes 16-17 = year (big-endian), 18 = month, 19 = day.
	// Returns null when the bytes are all zero (series master, not occurrence-specific).
	private getOccurrenceDateFromGlobalId(hexStr: string | undefined): moment.Moment | null {
		if (!hexStr || hexStr.length < 40 || hexStr.length % 2 || !/^[\da-f]+$/i.test(hexStr)) return null;
		const year = parseInt(hexStr.slice(32, 36), 16);
		const month = parseInt(hexStr.slice(36, 38), 16);
		const day = parseInt(hexStr.slice(38, 40), 16);
		if (!year || !month || !day) return null;
		const date = moment([year, month - 1, day]);
		return date.isValid() ? date : null;
	}

	// Try to extract the occurrence start date from the plain-text representation
	// that Outlook Classic puts in the DataTransfer when the user drags a calendar
	// event. The text contains a line like:
	//   English: "Start:   Wednesday, October 15, 2025 9:30 PM"
	//   French:  "Début :  mercredi 15 octobre 2025 21:30"
	// Returns null if no parseable date is found (caller should fall back to dialog).
	private parseDateFromDragText(text: string): moment.Moment | null {
		const match = text.match(/^[\t ]*(start|d[eé]but|begin)[\t ]*:[\t ]*(.+)$/im);
		if (!match) return null;
		const raw = match[2].trim().replace(/[\u00a0\u202f]/g, ' ');
		const french = /^d[eé]but$/i.test(match[1]);
		const englishFormats = ['dddd, MMMM D, YYYY h:mm A', 'dddd, MMMM D, YYYY HH:mm', 'dddd MMMM D, YYYY h:mm A', 'dddd MMMM D YYYY h:mm A', 'MMMM D, YYYY h:mm A', 'M/D/YYYY h:mm A', 'MM/DD/YYYY h:mm A', 'M/D/YYYY HH:mm', 'MM/DD/YYYY HH:mm'];
		const frenchFormats = ['dddd D MMMM YYYY HH:mm', 'dddd D MMMM YYYY H:mm', 'dddd D MMMM YYYY', 'D/M/YYYY HH:mm', 'DD/MM/YYYY HH:mm', 'D/M/YYYY H:mm'];
		const date = moment(raw, french ? frenchFormats : englishFormats, french ? 'fr' : 'en', true);
		if (date.isValid()) return date;
		const iso = moment(raw, moment.ISO_8601, true);
		return iso.isValid() ? iso : null;
	}

	// Return the nth (1-4, or 5 = last) weekday matching dayOfWeekBits within the
	// given month, at baseTime's time-of-day. Used for "MonthNth" recurrence patterns
	// (e.g. "the fourth Wednesday of every month"), where the day-of-month shifts from
	// month to month and cannot be found by simply adding months to the first occurrence.
	// Returns null for an invalid weekday mask or occurrence number.
	private nthWeekdayOfMonth(year: number, month0: number, dayOfWeekBits: number, n: number, baseTime: moment.Moment): moment.Moment | null {
		if (!Number.isInteger(n) || n < 1 || n > 5
			|| !Number.isInteger(dayOfWeekBits) || dayOfWeekBits < 1 || dayOfWeekBits > 127) return null;
		const matches: moment.Moment[] = [];
		const cursor = baseTime.clone().startOf('day').date(1).year(year).month(month0);
		const daysInMonth = cursor.daysInMonth();
		for (let d = 1; d <= daysInMonth; d++) {
			const day = cursor.clone().date(d);
			if (dayOfWeekBits & (1 << day.day())) matches.push(day);
		}
		if (matches.length === 0) return null;
		const picked = n === 5 ? matches[matches.length - 1] : matches[n - 1];
		if (!picked) return null;
		return picked.set({ hour: baseTime.hour(), minute: baseTime.minute(), second: baseTime.second(), millisecond: baseTime.millisecond() });
	}

	// Return the occurrence of a recurring series closest to `today`.
	// baseTime is apptStartWhole as a moment — the first occurrence with the
	// correct timezone. Using it as the anchor avoids the one-day-off error that
	// arises when startDate (always midnight UTC) is converted to local time.
	// Respects the series end date so past-ended or future series return the
	// closest valid occurrence rather than falling back to the series start.
	private findClosestOccurrence(apptRecur: AppointmentRecur, baseTime: moment.Moment, today: moment.Moment, distance?: (date: moment.Moment) => number): moment.Moment | null {
		try {
			const rp = apptRecur.recurrencePattern;
			const freq: number = rp.recurFrequency;
			const period: number = rp.period;
			if (!baseTime.isValid() || !today.isValid() || !Number.isInteger(period) || period < 1) return null;
			// Unsupported calendars must not silently use Gregorian arithmetic.
			if (![0, 1, 2, 9, 10, 11, 12].includes(rp.calendarType ?? 0)
				|| ![PatternType.Day, PatternType.Week, PatternType.Month, PatternType.MonthNth, PatternType.MonthEnd].includes(rp.patternType)) return null;
			if (rp.patternType === PatternType.MonthNth && !rp.patternTypeMonthNth) return null;

			const firstMidnight = baseTime.clone().startOf('day');
			const recurrenceDate = (n: number): moment.Moment => {
				const date = this.dateFromRecurMinutes(n).utc();
				return baseTime.isUTC() ? date : date.local(true);
			};
			// Recurrence dates encode local calendar dates, not UTC instants.
			const OUTLOOK_NO_END = 0x5AE980DF;
			const lastOccDate = rp.endDate && rp.endDate !== OUTLOOK_NO_END
				? recurrenceDate(rp.endDate) : null;
			if (lastOccDate && (!lastOccDate.isValid() || lastOccDate.isBefore(firstMidnight, 'day'))) return null;
			const inRange = (d: moment.Moment): boolean => d.isValid()
				&& !d.isBefore(firstMidnight, 'day') && (!lastOccDate || !d.isAfter(lastOccDate, 'day'));
			const dateKey = (d: moment.Moment): number => d.year() * 10000 + (d.month() + 1) * 100 + d.date();
			const excluded = new Set((rp.deletedInstanceDates ?? []).map(d =>
				dateKey(recurrenceDate(d))));
			const exceptions = apptRecur.exceptionInfo ?? [];
			for (const e of exceptions)
				excluded.add(dateKey(recurrenceDate(e.originalStartTime)));
			// Search past consecutive deletions as well as the adjacent periods.
			const radius = excluded.size + 1;
			const candidates: moment.Moment[] = [];
			const atTime = (d: moment.Moment): moment.Moment => d.set({
				hour: baseTime.hour(), minute: baseTime.minute(),
				second: baseTime.second(), millisecond: baseTime.millisecond()
			});
			const anchor = today.isBefore(firstMidnight) ? firstMidnight
				: (lastOccDate && today.isAfter(lastOccDate, 'day') ? lastOccDate : today);

			if (freq === 8202 && rp.patternType !== PatternType.Week) { // Daily: period is in minutes.
				if (period % 1440 !== 0) return null;
				const periodDays = period / 1440;
				const n = Math.floor(anchor.diff(firstMidnight, 'days') / periodDays);
				for (let i = Math.max(0, n - radius); i <= n + radius + 1; i++)
					candidates.push(baseTime.clone().add(i * periodDays, 'days'));
			} else if (freq === 8203 || (freq === 8202 && rp.patternType === PatternType.Week)) { // Weekly: use Outlook's week start, independent of locale.
				const dayBits = rp.patternTypeWeek?.dayOfWeekBits ?? (1 << baseTime.day());
				const firstDOW = rp.firstDOW ?? 0;
				if (!Number.isInteger(firstDOW) || firstDOW < 0 || firstDOW > 6
					|| !Number.isInteger(dayBits) || dayBits < 1 || dayBits > 127) return null;
				const firstWeek = firstMidnight.clone().subtract((firstMidnight.day() - firstDOW + 7) % 7, 'days');
				const n = Math.floor(anchor.diff(firstWeek, 'weeks') / period);
				for (let w = Math.max(0, n - radius); w <= n + radius + 1; w++) {
					const weekBase = firstWeek.clone().add(w * period, 'weeks');
					for (let d = 0; d < 7; d++)
						if (dayBits & (1 << ((firstDOW + d) % 7)))
							candidates.push(atTime(weekBase.clone().add(d, 'days')));
				}
			} else if (freq === 8204 || freq === 8205) { // Monthly/yearly: period is in months.
				const firstMonth = firstMidnight.clone().startOf('month');
				const n = Math.floor(anchor.diff(firstMonth, 'months') / period);
				const day = rp.patternTypeMonth?.day ?? baseTime.date();
				if (!Number.isInteger(day) || day < 1 || day > 31) return null;
				for (let i = Math.max(0, n - radius); i <= n + radius + 1; i++) {
					const month = firstMonth.clone().add(i * period, 'months');
					if (rp.patternType === PatternType.MonthNth && rp.patternTypeMonthNth) {
						const { dayOfWeekBits, n: nth } = rp.patternTypeMonthNth;
						const occ = this.nthWeekdayOfMonth(month.year(), month.month(), dayOfWeekBits, nth, baseTime);
						if (occ) candidates.push(occ);
					} else {
						const monthDay = rp.patternType === PatternType.MonthEnd ? month.daysInMonth() : Math.min(day, month.daysInMonth());
						candidates.push(atTime(month.date(monthDay)));
					}
				}
			} else {
				return null;
			}

			const valid = candidates.filter(c => inRange(c) && !excluded.has(dateKey(c)));
			// A moved exception may fall beyond the nominal first/last date. Its
			// original slot determines membership in the series.
			for (const e of exceptions) {
				const original = recurrenceDate(e.originalStartTime);
				const moved = recurrenceDate(e.startDateTime);
				if (inRange(original) && moved.isValid()) valid.push(moved);
			}
			if (valid.length === 0) return null;
			return valid.reduce((best, c) =>
				Math.abs(distance ? distance(c) : c.diff(today)) < Math.abs(distance ? distance(best) : best.diff(today)) ? c : best
			);
		} catch {
			return null;
		}
	}


	private parseMeetingDate(value: unknown): moment.Moment {
		return typeof value === 'string' && value.trim()
			? moment(value, [moment.ISO_8601, moment.RFC_2822], true) : moment.invalid();
	}

	private sanitizeNoteName(value: string): string {
		const forbidden = /[\u0000-\u001f\u007f<>:"/\\|?*]/g;
		const replacement = this.settings.invalidFilenameCharReplacement.replace(forbidden, '');
		let name = value.replace(forbidden, () => replacement).replace(/[ .]+$/, '').trim();
		if (!name || name === '.' || name === '..') name = 'Meeting';
		if (/^(con|prn|aux|nul|com[1-9¹²³]|lpt[1-9¹²³])(?:\.|$)/i.test(name)) name = '_' + name;
		// Leave room for the extension and avoid splitting a Unicode character.
		while (name.length > 240) name = Array.from(name).slice(0, -1).join('');
		return name.replace(/[ .]+$/, '') || 'Meeting';
	}

	private wallToInstant(wall: moment.Moment, offsetAt: (stamp: number) => number): moment.Moment {
		const stamp = wall.valueOf();
		const offsets = new Set([-36, 0, 36].map(h => offsetAt(stamp + h * 3600000)));
		const candidates = [...offsets].map(offset => stamp - offset * 60000).sort((a, b) => a - b);
		const exact = candidates.find(candidate => candidate + offsetAt(candidate) * 60000 === stamp);
		if (exact !== undefined) return moment.utc(exact); // First instance of a repeated wall time.
		// RFC 5545: a time in a spring gap uses the offset before the gap.
		const after = candidates.filter(candidate => candidate + offsetAt(candidate) * 60000 > stamp);
		if (!after.length) throw new TypeError('The meeting time cannot be resolved in its time zone.');
		return moment.utc(after.reduce((a, b) => a + offsetAt(a) * 60000 < b + offsetAt(b) * 60000 ? a : b));
	}

	private getMeetingClock(data?: MeetingFileData, zoneName?: string): MeetingClock {
		const rules = data?.apptTZDefRecur?.rules ?? [];
		const zoneStruct = data?.timeZoneStruct;
		if (rules.length || zoneStruct) {
			const offsetAt = (stamp: number): number => {
				const eligible = rules.filter(r => !r.start || Date.parse(r.start) <= stamp)
					.sort((a, b) => (a.start ? Date.parse(a.start) : -Infinity) - (b.start ? Date.parse(b.start) : -Infinity));
				const rule = eligible[eligible.length - 1] ?? zoneStruct ?? rules[0];
				if (!rule || ![rule.bias, rule.standardBias, rule.daylightBias].every(Number.isFinite)) throw new TypeError('The meeting contains invalid time zone data.');
				const standard = -rule.bias - rule.standardBias;
				const daylight = -rule.bias - rule.daylightBias;
				if (!rule.daylightDate.month || !rule.standardDate.month || standard === daylight) return standard;
				const year = new Date(stamp).getUTCFullYear();
				const transition = (t: OutlookZoneTransition, previousOffset: number): number => {
					const y = t.year || year;
					let day = t.day;
					if (!t.year) {
						const firstDOW = new Date(Date.UTC(y, t.month - 1, 1)).getUTCDay();
						day = 1 + (t.dayOfWeek - firstDOW + 7) % 7 + (t.day - 1) * 7;
						if (day > new Date(Date.UTC(y, t.month, 0)).getUTCDate()) day -= 7;
					}
					return Date.UTC(y, t.month - 1, day, t.hour, t.minute) - previousOffset * 60000;
				};
				const begins = transition(rule.daylightDate, standard);
				const ends = transition(rule.standardDate, daylight);
				return (begins < ends ? stamp >= begins && stamp < ends : stamp >= begins || stamp < ends) ? daylight : standard;
			};
			return {
				toWall: instant => moment.utc(instant.valueOf() + offsetAt(instant.valueOf()) * 60000),
				fromWall: wall => this.wallToInstant(wall, offsetAt)
			};
		}
		const name = zoneName ?? data?.apptTZDefRecur?.keyName ?? data?.timeZoneDesc;
		if (!name) return { toWall: instant => instant.clone().local().utc(true), fromWall: wall => wall.clone().local(true) };
		const zone = findIana(name, '001')[0] ?? name;
		let formatter: Intl.DateTimeFormat;
		try {
			formatter = new Intl.DateTimeFormat('en-US-u-ca-gregory-nu-latn', {
				timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit',
				hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23'
			});
		} catch { throw new TypeError('Unsupported meeting time zone: ' + name + '. Export the occurrence with its time zone definition.'); }
		const toWall = (instant: moment.Moment): moment.Moment => {
			const parts = Object.fromEntries(formatter.formatToParts(instant.toDate()).map(p => [p.type, p.value]));
			return moment.utc([+parts.year, +parts.month - 1, +parts.day, +parts.hour, +parts.minute, +parts.second, instant.millisecond()]);
		};
		const offsetAt = (stamp: number): number => (toWall(moment.utc(stamp)).valueOf() - stamp) / 60000;
		return { toWall, fromWall: wall => this.wallToInstant(wall, offsetAt) };
	}

	private getIcsClock(calendar: IcalComponent, property: IcalProperty, time: IcalTime): MeetingClock {
		const tzid = property.getParameter('tzid');
		if (time.isDate && tzid) throw new TypeError('An all-day calendar date must not have a time zone parameter.');
		if (time.zone === ICAL.Timezone.utcTimezone) return this.getMeetingClock(undefined, 'UTC');
		if (!tzid) return this.getMeetingClock();
		if (typeof tzid !== 'string') throw new TypeError('The calendar time zone is invalid.');
		const embedded = calendar.getTimeZoneByID(tzid);
		if (!embedded) return this.getMeetingClock(undefined, tzid);
		return {
			toWall: instant => {
				const t = ICAL.Time.fromJSDate(instant.toDate(), true).convertToZone(embedded);
				return moment.utc([t.year, t.month - 1, t.day, t.hour, t.minute, t.second]);
			},
			fromWall: wall => moment.utc(new ICAL.Time({
				year: wall.year(), month: wall.month() + 1, day: wall.date(),
				hour: wall.hour(), minute: wall.minute(), second: wall.second(), isDate: false
			}, embedded).toUnixTime() * 1000)
		};
	}

	private async correctIcsOccurrenceDate(fileData: MeetingFileData): Promise<boolean> {
		const occurrence = fileData.icsOccurrence;
		if (!occurrence) throw new TypeError('The recurring calendar file has no usable occurrence information.');
		const selected = await new Promise<string | null>(resolve => new OccurrenceDateModal(this.app, '', resolve).open());
		if (!selected) return false;
		const date = moment(selected, 'YYYY-MM-DD', 'en', true);
		if (!date.isValid()) return false;
		const { clock, startWall, calendarDays, seconds } = occurrence;
		const candidate = moment.utc([date.year(), date.month(), date.date(), startWall.hour(), startWall.minute(), startWall.second()]);
		const chosen = [-1, 0, 1].map(n => candidate.clone().add(n, 'days')).find(d => clock.fromWall(d).local().isSame(date, 'day'));
		if (!chosen) throw new TypeError('This occurrence date cannot be resolved in the meeting time zone.');
		fileData.apptStartWhole = clock.fromWall(chosen).toISOString();
		fileData.apptEndWhole = clock.fromWall(chosen.clone().add(calendarDays, 'days')).add(seconds, 'seconds').toISOString();
		return true;
	}
}

// Modal shown when the specific occurrence date cannot be determined from the .msg file.
// The user can confirm or correct the pre-filled date before the note is created.
class OccurrenceDateModal extends Modal {
	private dateStr: string;
	private readonly onSubmit: (date: string | null) => void;
	private resolved = false;

	constructor(app: App, suggestedDate: string, onSubmit: (date: string | null) => void) {
		super(app);
		this.dateStr = suggestedDate;
		this.onSubmit = onSubmit;
	}

	private resolve(date: string | null): void {
		if (this.resolved) return;
		this.resolved = true;
		this.onSubmit(date);
	}

	onOpen(): void {
		const { contentEl } = this;
		new Setting(contentEl).setName('Confirm occurrence date').setHeading();
		contentEl.createEl('p', {
			text: 'Outlook did not provide a usable occurrence date. '
				+ (this.dateStr
					? 'The field below is pre-filled with the occurrence nearest to today. '
					: 'An occurrence could not be calculated for this series. ')
				+ 'Confirm or enter the date shown in your calendar.'
		});

		new Setting(contentEl)
			.setName('Event date')
			.addText(text => {
				text.inputEl.type = 'date';
				text.setValue(this.dateStr);
				text.onChange(value => { this.dateStr = value; });
				text.inputEl.addEventListener('keydown', (e) => {
					if (e.key === 'Enter') { this.resolve(this.dateStr); this.close(); }
				});
			});

		new Setting(contentEl)
			.addButton(btn => btn
				.setButtonText('Create note')
				.setCta()
				.onClick(() => { this.resolve(this.dateStr); this.close(); }))
			.addButton(btn => btn
				.setButtonText('Cancel')
				.onClick(() => { this.resolve(null); this.close(); }));
	}

	onClose(): void {
		this.contentEl.empty();
		this.resolve(null); // no-op if already resolved via a button
	}
}

class OutlookMeetingNotesSettingTab extends PluginSettingTab {
	plugin: OutlookMeetingNotes;

	constructor(app: App, plugin: OutlookMeetingNotes) {
		super(app, plugin);
		this.plugin = plugin;
	}

	// Declarative metadata so these settings are found by Obsidian's settings
	// search (available since 1.13.0). Rendering is still handled by display()
	// below, since the imperative API gives finer control over the template
	// text area and the documentation/donation links.
	getSettingDefinitions(): SettingDefinitionItem[] {
		return [
			{
				name: 'Folder location',
				desc: 'Notes will be created in this folder.',
				control: { type: 'text', key: 'notesFolder' },
			},
			{
				name: 'Filename pattern',
				desc: 'This pattern will be used to name new notes.',
				control: { type: 'text', key: 'fileNamePattern' },
			},
			{
				name: 'Invalid character substitute',
				desc: 'This character (or string) will be used in place of any invalid characters for new note filenames.',
				control: { type: 'text', key: 'invalidFilenameCharReplacement' },
			},
			{
				name: 'Template',
				desc: 'This template will be used for new notes.',
				control: { type: 'textarea', key: 'notesTemplate' },
			},
		];
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		new Setting(containerEl)
			.setName('Folder location')
			.setDesc('Notes will be created in this folder.')
			.addText(text => text
				.setPlaceholder('Example: folder 1/subfolder 2')
				.setValue(this.plugin.settings.notesFolder)
				.onChange(async (value) => {
					this.plugin.settings.notesFolder = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Filename pattern')
			.setDesc('This pattern will be used to name new notes.')
			.addText(text => text
				.setPlaceholder('Default: ' + OutlookMeetingNotesDefaultFilenamePattern)
				.setValue(this.plugin.settings.fileNamePattern)
				.onChange(async (value) => {
					if (value == '') {
						this.plugin.settings.fileNamePattern = OutlookMeetingNotesDefaultFilenamePattern;
					} else {
						this.plugin.settings.fileNamePattern = value;
					}
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Invalid character substitute')
			.setDesc('This character (or string) will be used in place of any invalid characters for new note filenames.')
			.addText(text => text
				.setPlaceholder('Example: _')
				.setValue(this.plugin.settings.invalidFilenameCharReplacement)
				.onChange(async (value) => {
					this.plugin.settings.invalidFilenameCharReplacement = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Template')
			.setDesc('This template will be used for new notes.')
			.addTextArea(text => text
				.setPlaceholder('Default: ' + OutlookMeetingNotesDefaultFilenamePattern)
				.setValue(this.plugin.settings.notesTemplate)
				.onChange(async (value) => {
					this.plugin.settings.notesTemplate = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setDesc(createFragment(df => {
				df.appendText('For more information about filename patterns and the syntax for templates, see the ');
				df.createEl('a', {
					text: 'Documentation',
					href: 'https://github.com/nathanstorm689/outlook-event-notes#readme',
					attr: { target: '_blank', rel: 'noopener' }
				});
				df.appendText('.');
			}));

		new Setting(containerEl)
			.setDesc(createFragment(df => {
				df.appendText('If this plugin saves you time, consider ');
				df.createEl('a', {
					text: 'Buying me a coffee',
					href: 'https://buymeacoffee.com/nathanstorm',
					attr: { target: '_blank', rel: 'noopener' }
				});
				df.appendText(' ☕');
			}));
	}
}
