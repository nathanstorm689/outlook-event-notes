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
			const { vault } = this.app;

			// Get the file data from MsgReader
			const origFileData = msg.getFileData();

			// Check if we got a suitable meeting
			if (origFileData.dataType != 'msg') {
				throw new TypeError('Outlook Event Notes cannot process the file. '
					+ 'MsgReader did not parse the file as valid msg format.');
			} else if (origFileData.messageClass != 'IPM.Appointment') {
				throw new TypeError('Outlook Event Notes cannot process the file. '
					+ 'It is a valid msg file but not an appointment or meeting.');
			}

			this.addHelperFunctions(origFileData);

			// Add helper field for the current date and time
			let fileData = origFileData as any;
			fileData.helper_currentDT = moment().format();

			this.ensureBodyField(fileData);

			// For recurring events, apptStartWhole may contain the first occurrence's date
			// rather than the specific dragged occurrence. Correct it.
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
				// File already exists
				new Notice(meetingNoteFile.basename + ' already exists: opening it');
			}
			else {
				if (targetFolderPath !== '' && vault.getFolderByPath(targetFolderPath) == null) {
					await vault.createFolder(targetFolderPath);
				}
				const mustacheOutput = this.renderTemplate(
					this.settings.notesTemplate,
					fileData);
				meetingNoteFile = await vault.create(filePath, mustacheOutput);
				new Notice('New file created: ' + meetingNoteFile.basename);
			}
			const openInNewTab = false;
			this.app.workspace.getLeaf(openInNewTab).openFile(meetingNoteFile);
			// @ts-ignore: Property 'internalPlugins' does not exist on type 'App'.
			const fe = this.app.internalPlugins.getEnabledPluginById("file-explorer");
			if (fe) { fe.revealInFolder(meetingNoteFile); }
		} catch (ee: unknown) {
			if (ee instanceof Error) { new Notice('Error (' + ee.name + '):\n' + ee.message); }
			throw ee;
		}
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
	async handleDropEvent(dropevt: DragEvent) {
		if (dropevt.dataTransfer == null) {
			throw new ReferenceError('Outlook Event Notes cannot handle the DragEvent. The event had a null '
				+ 'dataTransfer property, which should never happen when dispatched by the browser, according '
				+ 'to https://developer.mozilla.org/en-US/docs/Web/API/DragEvent/dataTransfer');
		} else {
			const droppedFiles = dropevt.dataTransfer.files
			if (droppedFiles.length != 1) {
				new Notice('Outlook Event Notes can only handle one meeting being dropped onto the ribbon icon');
			}
			else {
				const droppedFile = droppedFiles[0];

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
				}
				fr.readAsArrayBuffer(droppedFile)
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
		const firstOccDate = this.dateFromRecurMinutes(rp.startDate);
		if (apptStart.local().format('YYYY-MM-DD') !== firstOccDate.local().format('YYYY-MM-DD')) return true;

		const timeOffset: number = apptRecur.startTimeOffset ?? 0;
		let corrected: moment.Moment | null = null;

		// Try PidLidGlobalObjectId first — most reliable source for native Outlook events.
		const occDateFromId = this.getOccurrenceDateFromGlobalId(fileData.globalAppointmentID);
		if (occDateFromId) {
			corrected = occDateFromId.startOf('day').add(timeOffset, 'minutes');
		} else {
			// The .msg file does not encode the specific occurrence date (common for
			// Google Calendar / third-party events synced to Outlook). Ask the user.
			const suggested = this.findClosestOccurrence(apptRecur, moment());
			const suggestedStr = suggested
				? suggested.local().format('YYYY-MM-DD')
				: apptStart.local().format('YYYY-MM-DD');

			const userDateStr = await new Promise<string | null>((resolve) => {
				new OccurrenceDateModal(this.app, suggestedStr, resolve).open();
			});

			if (!userDateStr) return false; // user cancelled

			const userDate = moment(userDateStr, 'YYYY-MM-DD', true);
			if (!userDate.isValid()) return false;
			corrected = userDate.startOf('day').add(timeOffset, 'minutes');
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
	private findClosestOccurrence(apptRecur: any, today: moment.Moment): moment.Moment | null {
		try {
			const rp = apptRecur.recurrencePattern;
			const timeOffset: number = apptRecur.startTimeOffset ?? 0;

			const firstMidnight = this.dateFromRecurMinutes(rp.startDate);
			const firstOcc = firstMidnight.clone().startOf('day').add(timeOffset, 'minutes');

			const candidates: moment.Moment[] = [];
			const freq: number = rp.recurFrequency;
			const period: number = rp.period;

			if (freq === 8202) { // Daily (period is in minutes)
				const periodDays = Math.max(1, Math.round(period / 1440));
				const n = Math.round(today.diff(firstMidnight, 'days') / periodDays);
				for (let i = Math.max(0, n - 1); i <= n + 2; i++)
					candidates.push(firstOcc.clone().add(i * periodDays, 'days'));
			} else if (freq === 8203) { // Weekly (period is in weeks)
				const dayBits: number = rp.patternTypeWeek?.dayOfWeekBits ?? (1 << firstMidnight.day());
				const firstWeekSun = firstMidnight.clone().startOf('week');
				const n = Math.round(today.diff(firstWeekSun, 'weeks') / period);
				for (let w = Math.max(0, n - 1); w <= n + 2; w++) {
					const weekBase = firstWeekSun.clone().add(w * period, 'weeks');
					for (let d = 0; d < 7; d++)
						if (dayBits & (1 << d))
							candidates.push(weekBase.clone().add(d, 'days').add(timeOffset, 'minutes'));
				}
			} else if (freq === 8204) { // Monthly (period is in months)
				const n = Math.round(Math.max(0, today.diff(firstMidnight, 'months')) / period);
				for (let i = Math.max(0, n - 1); i <= n + 2; i++)
					candidates.push(firstOcc.clone().add(i * period, 'months'));
			} else if (freq === 8205) { // Yearly
				const n = Math.max(0, today.diff(firstMidnight, 'years'));
				for (let i = Math.max(0, n - 1); i <= n + 2; i++)
					candidates.push(firstOcc.clone().add(i, 'years'));
			} else {
				return null;
			}

			const valid = candidates.filter(c => !c.isBefore(firstOcc, 'day'));
			if (valid.length === 0) return firstOcc;
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
