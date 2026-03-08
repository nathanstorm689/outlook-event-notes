import { App, displayTooltip, Modal, Notice, Plugin, PluginSettingTab, Setting, TooltipPlacement, normalizePath } from 'obsidian';
import MsgReader from '@kenjiuno/msgreader';
import proxyData from 'mustache-validator';
import Mustache from 'mustache';
import moment from 'moment';

const OutlookMeetingNotesDefaultFilenamePattern =
	'{{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD HH.mm{{/helper_dateFormat}} {{subject}}';

const OutlookMeetingNotesDefaultTemplate = `---
title: {{subject}}
subtitle: meeting notes
date: {{#helper_dateFormat}}{{apptStartWhole}}|L LT{{/helper_dateFormat}}
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

	async createMeetingNote(msg: MsgReader) {
		try {
			const origFileData = msg.getFileData();
			if (origFileData.dataType != 'msg') {
				throw new TypeError('Outlook Event Notes cannot process the file. '
					+ 'MsgReader did not parse the file as valid msg format.');
			} else if (origFileData.messageClass != 'IPM.Appointment') {
				throw new TypeError('Outlook Event Notes cannot process the file. '
					+ 'It is a valid msg file but not an appointment or meeting.');
			}
			await this.createNoteFromFileData(origFileData as any);
		} catch (ee: unknown) {
			if (ee instanceof Error) { new Notice('Error (' + ee.name + '):\n' + ee.message); }
			throw ee;
		}
	}

	// Shared note-creation logic used by both the .msg path and the .ics path.
	// fileData must have: subject, apptStartWhole, apptEndWhole, apptLocation,
	// body/bodyText/bodyHtml, recipients[], apptRecur (null = skip date correction).
	private async createNoteFromFileData(origFileData: any): Promise<void> {
		const { vault } = this.app;

		this.addHelperFunctions(origFileData);
		let fileData = origFileData;
		fileData.helper_currentDT = moment().format();

		this.ensureBodyField(fileData);

		// For recurring .msg events, apptStartWhole may carry only the series start date.
		// correctRecurringOccurrenceDate fixes it (or asks the user).
		// For .ics files apptRecur is null, so this returns true immediately.
		const dateOk = await this.correctRecurringOccurrenceDate(fileData);
		if (!dateOk) return; // user cancelled the date dialog

		let targetFolderPath = (this.settings.notesFolder ?? '').trim();
		if (targetFolderPath === '' || targetFolderPath === '/') { targetFolderPath = ''; }
		else { targetFolderPath = normalizePath(targetFolderPath); }
		const fileNameEscape = {
			escape: (str: string): string => {
				return str.replaceAll('/', this.settings.invalidFilenameCharReplacement);
			}
		}
		const fileNameMustache = Mustache.render(
			this.settings.fileNamePattern,
			proxyData(fileData),
			undefined,
			fileNameEscape)
			.replaceAll(/[*\"\\<>:|?]/g, this.settings.invalidFilenameCharReplacement);
		const folderPrefix = targetFolderPath === '' ? '' : targetFolderPath + '/';
		const filePath = normalizePath(folderPrefix + fileNameMustache + '.md');
		let meetingNoteFile = vault.getFileByPath(filePath);
		if (meetingNoteFile) {
			new Notice(meetingNoteFile.basename + ' already exists: opening it');
		} else {
			if (targetFolderPath !== '' && vault.getFolderByPath(targetFolderPath) == null) {
				await vault.createFolder(targetFolderPath);
			}
			const mustacheOutput = this.renderTemplate(this.settings.notesTemplate, fileData);
			meetingNoteFile = await vault.create(filePath, mustacheOutput);
			new Notice('New file created: ' + meetingNoteFile.basename);
		}
		const openInNewTab = false;
		this.app.workspace.getLeaf(openInNewTab).openFile(meetingNoteFile);
		// @ts-ignore: Property 'internalPlugins' does not exist on type 'App'.
		const fe = this.app.internalPlugins.getEnabledPluginById("file-explorer");
		if (fe) { fe.revealInFolder(meetingNoteFile); }
	}

	private ensureBodyField(fileData: any): void {
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
	// Drag a meeting directly from the Outlook Calendar view to get an .ics file
	// whose DTSTART is always the exact occurrence date (no dialog needed).
	async handleDropEvent(dropevt: DragEvent) {
		if (dropevt.dataTransfer == null) {
			throw new ReferenceError('Outlook Event Notes cannot handle the DragEvent. The event had a null '
				+ 'dataTransfer property, which should never happen when dispatched by the browser, according '
				+ 'to https://developer.mozilla.org/en-US/docs/Web/API/DragEvent/dataTransfer');
		} else {
			const droppedFiles = dropevt.dataTransfer.files;
			if (droppedFiles.length != 1) {
				new Notice('Outlook Event Notes can only handle one meeting being dropped onto the ribbon icon');
			} else {
				const droppedFile = droppedFiles[0];
				const isIcs = droppedFile.name.toLowerCase().endsWith('.ics')
					|| droppedFile.type === 'text/calendar';

				if (isIcs) {
					// iCalendar file: read as text and parse.
					// DTSTART is always the correct occurrence date → no dialog needed.
					const fr = new FileReader();
					fr.onload = async () => {
						try {
							const fileData = this.parseIcsFile(fr.result as string);
							await this.createNoteFromFileData(fileData);
						} catch (ee: unknown) {
							if (ee instanceof Error) { new Notice('Error (' + ee.name + '):\n' + ee.message); }
						}
					};
					fr.readAsText(droppedFile, 'utf-8');
				} else {
					// Outlook .msg binary file
					const fr = new FileReader();
					fr.onload = async () => {
						if (fr.result == null) {
							throw new ReferenceError('Outlook Event Notes cannot handle the DragEvent. The FileReader had '
								+ 'a null result property, which should not be possible.');
						} else if (!(fr.result instanceof ArrayBuffer)) {
							throw new TypeError('Outlook Event Notes cannot handle the DragEvent. The FileReader result '
								+ 'property was not an ArrayBuffer, which should be impossible.');
						} else {
							const msgRdr = new MsgReader(fr.result);
							await this.createMeetingNote(msgRdr);
						}
					};
					fr.readAsArrayBuffer(droppedFile);
				}
			}
		}
	}

	private ribbonIconEl: HTMLElement;

	async onload() {
		await this.loadSettings();

		const tooltipMessage = 'Outlook Event Notes: Drag and drop a meeting onto this icon from Outlook (or a .msg file) to create a meeting note.';

		this.ribbonIconEl = this.addRibbonIcon('calendar-clock', tooltipMessage, () => { });

		this.ribbonIconEl.addEventListener('dragenter', () => {
			this.ribbonIconEl.toggleClass('is-being-dragged-over', true);
			const ttPosition = this.ribbonIconEl.getAttribute('data-tooltip-position') as TooltipPlacement;
			const ttDelay = this.ribbonIconEl.getAttribute('data-tooltip-delay');
			if (ttPosition != null && ttDelay != null) {
				displayTooltip(this.ribbonIconEl, tooltipMessage, { placement: ttPosition, delay: Number(ttDelay) });
			} else {
				displayTooltip(this.ribbonIconEl, tooltipMessage);
			}
		});
		this.ribbonIconEl.addEventListener('dragleave', () => {
			this.ribbonIconEl.toggleClass('is-being-dragged-over', false);
			const tooltip = document.getElementsByClassName('tooltip')[0]
			if (tooltip) { tooltip.remove(); }
		});
		this.ribbonIconEl.addEventListener('dragover', (dragevt) => {
			dragevt.preventDefault();
			if (dragevt.dataTransfer != null) { dragevt.dataTransfer.dropEffect = 'copy'; }
		});
		this.ribbonIconEl.addEventListener('drop', (dropevt) => {
			this.ribbonIconEl.toggleClass('is-being-dragged-over', false);
			dropevt.preventDefault();
			this.handleDropEvent(dropevt);
		});
		this.ribbonIconEl.addClass('outlook-event-notes-icon');

		this.addSettingTab(new OutlookMeetingNotesSettingTab(this.app, this));
	}

	onunload() { }

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	addHelperFunctions(hash: any): any {
		const helperFunctions = {
			firstWord: () => {
				return function (words: string, render: any) {
					return render(words).replace(/\W.*$/, '');
				}
			},
			dateFormat: () => {
				return function (datetime_format: string, render: any) {
					const parts = datetime_format.split('|');
					const rawValue = render(parts[0]).trim();
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
		for (let property in hash) {
			if (hash[property] instanceof Array) {
				for (let subproperty in hash[property]) {
					if (hash[property][subproperty] instanceof Object) {
						for (func in helperFunctions) {
							hash[property][subproperty]['helper_' + func] = helperFunctions[func];
						}
					}
				}
			}
		}

		return hash;
	}

	// Parse template into YAML and markdown sections to use different escaping for each
	renderTemplate(template: string, hash: any): string {
		// Matches '---' frontmatter block at the start of the string
		const templateYAMLMatch = template.match(/^---(\r\n?|\n).*?(\r\n?|\n)---($|\r\n?|\n)/s);
		const templateMD = templateYAMLMatch ? template.substring(templateYAMLMatch[0].length) : template;

		let output = ''

		if (templateYAMLMatch) {
			const sanitizeYamlValue = (value: string): string => value.replace(/[><*]/g, '');
			const mustacheYAMLOptions = {
				escape: (str: string): string => {
					const sanitized = sanitizeYamlValue(str);
					const found = sanitized.match(/\r\n?|\n/);
					if (found) {
						return '|\n' + '  ' + sanitized.replaceAll(/\r\n?|\n/g, '\n  ');
					} else if (sanitized.match(/[:#\[\]\{\},]/)) {
						return '"' + sanitized.replaceAll(/["\\]/g, '\$&') + '"';
					}
					else return sanitized;
				}
			}

			output = output + Mustache.render(
				templateYAMLMatch[0],
				proxyData(hash),
				undefined,
				mustacheYAMLOptions);
		}

		if (templateMD) {
			const mustacheMDOptions = {
				escape: (str: string): string => {
					return str.replaceAll(/[\\\`\*\_\[\]\{\}\<\>\(\)\#\!\|\^]/g, '\\$&')
						.replaceAll('%%', '\\%\\%')
						.replaceAll('~~', '\\~\\~')
						.replaceAll('==', '\\=\\=');
				}
			}

			output = output + Mustache.render(
				templateMD,
				proxyData(hash),
				undefined,
				mustacheMDOptions);
		}

		return output;
	}


	// Parse an iCalendar (.ics) file and return a fileData object compatible with
	// createNoteFromFileData. DTSTART in .ics files is always the correct occurrence
	// date (Outlook sets it to the specific occurrence when dragging from calendar view),
	// so apptRecur is set to null to skip the recurring-event correction dialog.
	private parseIcsFile(content: string): any {
		// RFC 5545 §3.1: unfold continuation lines (lines starting with whitespace)
		const text = content
			.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
			.replace(/\n[ \t]/g, '');

		// Find first VEVENT block
		const m = text.match(/BEGIN:VEVENT([\s\S]*?)END:VEVENT/i);
		if (!m) throw new TypeError('The .ics file does not contain a VEVENT block.');
		const block = m[1];

		// Parse all property lines: NAME or NAME;PARAM=VAL:value
		const propMap = new Map<string, Array<{ value: string; params: Record<string, string> }>>();
		for (const line of block.split('\n')) {
			const colon = line.indexOf(':');
			if (colon === -1) continue;
			const keyFull = line.slice(0, colon);
			const value = line.slice(colon + 1).trimEnd();
			const semi = keyFull.indexOf(';');
			const name = (semi === -1 ? keyFull : keyFull.slice(0, semi)).toUpperCase();
			const paramStr = semi === -1 ? '' : keyFull.slice(semi + 1);
			const params: Record<string, string> = {};
			for (const seg of paramStr.split(';').filter(Boolean)) {
				const eq = seg.indexOf('=');
				if (eq !== -1) {
					params[seg.slice(0, eq).toUpperCase()] =
						seg.slice(eq + 1).replace(/^"(.*)"$/, '$1');
				}
			}
			if (!propMap.has(name)) propMap.set(name, []);
			propMap.get(name)!.push({ value, params });
		}

		const get = (name: string) => propMap.get(name)?.[0];

		// Unescape RFC 5545 text values
		const unescape = (s: string) =>
			s.replace(/\\n/gi, '\n').replace(/\\,/g, ',')
				.replace(/\\;/g, ';').replace(/\\\\/g, '\\');

		// Parse a date/time property to a moment.
		// UTC times end with Z; TZID times are parsed as local (assumes the
		// computer's timezone matches the event's timezone, which is the common case).
		const parseDT = (prop: { value: string; params: Record<string, string> } | undefined)
			: moment.Moment | undefined => {
			if (!prop) return undefined;
			const v = prop.value.replace(/\s/g, '');
			if (v.endsWith('Z')) {
				const clean = v.slice(0, -1);
				return moment.utc(clean, clean.includes('T') ? 'YYYYMMDDTHHmmss' : 'YYYYMMDD');
			}
			return moment(v, v.includes('T') ? 'YYYYMMDDTHHmmss' : 'YYYYMMDD');
		};

		// Parse attendees (ATTENDEE;CN=...:mailto:email)
		const recipients: Array<{ name: string; email: string }> = [];
		for (const att of propMap.get('ATTENDEE') ?? []) {
			const emailM = att.value.match(/mailto:(.+)/i);
			if (!emailM) continue;
			const email = emailM[1].trim();
			const cn = att.params['CN'];
			recipients.push({ name: cn ? unescape(cn) : email, email });
		}

		const summaryProp = get('SUMMARY');
		const dtstart = get('DTSTART');
		const dtend = get('DTEND');
		const locationProp = get('LOCATION');
		const descProp = get('DESCRIPTION');

		if (!summaryProp) throw new TypeError('The .ics file is missing a SUMMARY (event title).');
		if (!dtstart) throw new TypeError('The .ics file is missing a DTSTART (start date).');

		return {
			dataType: 'msg',                  // satisfies createMeetingNote validation path
			messageClass: 'IPM.Appointment',
			subject: unescape(summaryProp.value),
			apptStartWhole: parseDT(dtstart)?.toISOString(),
			apptEndWhole: parseDT(dtend)?.toISOString(),
			apptLocation: locationProp ? unescape(locationProp.value) : '',
			body: descProp ? unescape(descProp.value) : '',
			recipients,
			apptRecur: null,  // DTSTART already has the correct occurrence date
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
	private async correctRecurringOccurrenceDate(fileData: any): Promise<boolean> {
		const apptRecur = fileData.apptRecur;
		if (!apptRecur?.recurrencePattern) return true;

		const rp = apptRecur.recurrencePattern;
		const apptStart = moment(fileData.apptStartWhole);
		if (!apptStart.isValid()) return true;

		// If apptStartWhole already differs from the series start, Outlook already
		// provided the correct occurrence date (exception object). Leave it alone.
		// We compare by difference in hours rather than by local calendar date
		// because apptStartWhole and startDate can be on different calendar days
		// in local time (e.g. apptStartWhole = 23:00 UTC = 01:00 the next day in
		// UTC+2). A gap ≥ 24 h means they represent different occurrences.
		const firstOccDate = this.dateFromRecurMinutes(rp.startDate);
		if (Math.abs(apptStart.diff(firstOccDate, 'hours')) >= 24) return true;

		let corrected: moment.Moment | null = null;

		// Helper: given a chosen local date, produce the corrected occurrence moment.
		// We preserve the original time-of-day from apptStartWhole (in local timezone)
		// and only swap the calendar date, so the resulting UTC value is correct
		// regardless of whether the event was created in a different timezone.
		const withDate = (d: moment.Moment): moment.Moment =>
			apptStart.clone().year(d.year()).month(d.month()).date(d.date());

		// Try PidLidGlobalObjectId first — most reliable source for native Outlook events.
		const occDateFromId = this.getOccurrenceDateFromGlobalId(fileData.globalAppointmentID);
		if (occDateFromId) {
			corrected = withDate(occDateFromId);
		} else {
			// The .msg file does not encode the specific occurrence date (common for
			// Google Calendar / third-party events synced to Outlook). Ask the user.
			// Pre-fill with the series start date in local time — it won't be the exact
			// occurrence for non-first occurrences, but it gives the right time context
			// and the user can correct it to the actual date they see in Outlook.
			const suggestedStr = apptStart.local().format('YYYY-MM-DD');

			const userDateStr = await new Promise<string | null>((resolve) => {
				new OccurrenceDateModal(this.app, suggestedStr, resolve).open();
			});

			if (!userDateStr) return false; // user cancelled

			const userDate = moment(userDateStr, 'YYYY-MM-DD', true);
			if (!userDate.isValid()) return false;
			corrected = withDate(userDate);
		}

		if (!corrected) return true;

		const endStart = moment(fileData.apptEndWhole);
		if (endStart.isValid()) {
			const duration = endStart.diff(apptStart, 'minutes');
			fileData.apptEndWhole = corrected.clone().add(duration, 'minutes').toISOString();
		}
		fileData.apptStartWhole = corrected.toISOString();
		return true;
	}

	// Parse the occurrence date from PidLidGlobalObjectId (as hex string).
	// Bytes 16-17 = year (big-endian), 18 = month, 19 = day.
	// Returns null when the bytes are all zero (series master, not occurrence-specific).
	private getOccurrenceDateFromGlobalId(hexStr: string | undefined): moment.Moment | null {
		if (!hexStr || hexStr.length < 40) return null;
		const year = (parseInt(hexStr.substring(32, 34), 16) << 8) | parseInt(hexStr.substring(34, 36), 16);
		const month = parseInt(hexStr.substring(36, 38), 16);
		const day = parseInt(hexStr.substring(38, 40), 16);
		if (year === 0 || month === 0 || day === 0) return null;
		return moment({ year, month: month - 1, day }); // month is 0-indexed in moment
	}

	// Return the occurrence of a recurring series closest to `today`.
	// baseTime is apptStartWhole as a moment — the first occurrence with the
	// correct timezone. Using it as the anchor avoids the one-day-off error that
	// arises when startDate (always midnight UTC) is converted to local time.
	private findClosestOccurrence(apptRecur: any, baseTime: moment.Moment, today: moment.Moment): moment.Moment | null {
		try {
			const rp = apptRecur.recurrencePattern;

			// Local-time midnight of baseTime, used for day-level period arithmetic.
			const firstMidnight = baseTime.clone().startOf('day');

			const candidates: moment.Moment[] = [];
			const freq: number = rp.recurFrequency;
			const period: number = rp.period;

			if (freq === 8202) { // Daily (period is in minutes)
				const periodDays = Math.max(1, Math.round(period / 1440));
				const n = Math.round(today.diff(firstMidnight, 'days') / periodDays);
				for (let i = Math.max(0, n - 1); i <= n + 2; i++)
					candidates.push(baseTime.clone().add(i * periodDays, 'days'));
			} else if (freq === 8203) { // Weekly (period is in weeks)
				const dayBits: number = rp.patternTypeWeek?.dayOfWeekBits ?? (1 << baseTime.day());
				const firstWeekSun = firstMidnight.clone().startOf('week');
				const n = Math.round(today.diff(firstWeekSun, 'weeks') / period);
				for (let w = Math.max(0, n - 1); w <= n + 2; w++) {
					const weekBase = firstWeekSun.clone().add(w * period, 'weeks');
					for (let d = 0; d < 7; d++)
						if (dayBits & (1 << d))
							candidates.push(weekBase.clone().add(d, 'days')
								.add(baseTime.hours() * 60 + baseTime.minutes(), 'minutes'));
				}
			} else if (freq === 8204) { // Monthly (period is in months)
				const n = Math.round(Math.max(0, today.diff(firstMidnight, 'months')) / period);
				for (let i = Math.max(0, n - 1); i <= n + 2; i++)
					candidates.push(baseTime.clone().add(i * period, 'months'));
			} else if (freq === 8205) { // Yearly
				const n = Math.max(0, today.diff(firstMidnight, 'years'));
				for (let i = Math.max(0, n - 1); i <= n + 2; i++)
					candidates.push(baseTime.clone().add(i, 'years'));
			} else {
				return null;
			}

			const valid = candidates.filter(c => !c.isBefore(baseTime, 'day'));
			if (valid.length === 0) return baseTime;
			return valid.reduce((best, c) =>
				Math.abs(c.diff(today)) < Math.abs(best.diff(today)) ? c : best
			);
		} catch {
			return null;
		}
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
		contentEl.createEl('h3', { text: 'Confirm occurrence date' });
		contentEl.createEl('p', {
			text: 'The date of this recurring event occurrence could not be determined automatically. '
				+ 'Please confirm or correct the date (YYYY-MM-DD):'
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
			.setDesc((() => {
				const df = document.createDocumentFragment();
				df.appendChild(document.createTextNode(
					'For more information about filename patterns and the syntax for templates, see the '
				));
				const link = document.createElement('a');
				link.href = 'https://github.com/nathanstorm689/outlook-meeting-notes-plus#readme';
				link.target = '_blank';
				link.rel = 'noopener';
				link.textContent = 'documentation';
				df.appendChild(link);
				df.appendChild(document.createTextNode('.'))
				return df;
			})())

	}
}
