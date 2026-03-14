import { App, ButtonComponent, Modal, TextComponent, displayTooltip, Editor, MarkdownView, Notice, Plugin, PluginSettingTab, Setting, TooltipPlacement, EventRef, normalizePath } from 'obsidian';
import MsgReader from '@kenjiuno/msgreader';
import proxyData from 'mustache-validator';
import Mustache from 'mustache';
import moment from 'moment';

const OutlookMeetingNotesDefaultFilenamePattern =
	'{{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD_HH-mm_ss{{/helper_dateFormat}} {{subject}}';

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

	//TODO: Add functionality to export meeting notes nicely

	async createMeetingNote(msg: MsgReader) {
		try {
			const { vault } = this.app;

			// Get the file data from MsgReader
			const origFileData = msg.getFileData();

			// Check if we got a suitable meeting
			if (origFileData.dataType != 'msg') {
				throw new TypeError('Outlook Meeting Notes cannot process the file. '
					+ 'MsgReader did not parse the file as valid msg format.');
			} else if (origFileData.messageClass != 'IPM.Appointment') {
				throw new TypeError('Outlook Meeting Notes cannot process the file. '
					+ 'It is a valid msg file but not an appointment or meeting.');
			}

			this.addHelperFunctions(origFileData);

			// Add helper field for the current date and time
			let fileData = origFileData as any;
			fileData.helper_currentDT = moment().format();

			this.ensureBodyField(fileData);
			this.ensureDefaultFields(fileData);

			if (this.isRecurringAppointment(fileData)) {
				const occurrenceDate = await this.getRecurringOccurrenceDate();
				if (occurrenceDate === null) {
					new Notice('Meeting note creation cancelled.');
					return;
				}
				this.applyRecurringOccurrenceDate(fileData, occurrenceDate);
			}

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
			// TODO: Handle errors reasonably -- differently between msg missing elements and errors creating file
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

	private ensureDefaultFields(fileData: any): void {
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


	private isRecurringAppointment(fileData: any): boolean {
		const keys = Object.keys(fileData ?? {});
		const boolClues = new Set(['isrecurring', 'isrecurringmeeting', 'recurring']);
		for (const key of keys) {
			const lower = key.toLowerCase();
			if (boolClues.has(lower)) {
				if ((fileData as Record<string, unknown>)[key] === true) {
					return true;
				}
			}
		}
		const hintSubstrings = ['apptrecur', 'appointmentrecur', 'recurrence', 'recurrencerule', 'recurrencepattern', 'recurrenceinfo', 'recurrencestate', 'recurrencetype', 'recurringmaster', 'apptimezonedefrecur'];
		for (const key of keys) {
			const lower = key.toLowerCase();
			for (const hint of hintSubstrings) {
				if (lower.includes(hint)) {
					const value = (fileData as Record<string, unknown>)[key];
					if (typeof value === 'boolean') { return value; }
					if (value != null && value !== '') { return true; }
				}
			}
		}
		if (typeof fileData.messageClass === 'string') {
			const lowerClass = fileData.messageClass.toLowerCase();
			if (lowerClass.includes('recurring') || lowerClass.includes('exception') || lowerClass.includes('occurrence')) {
				return true;
			}
		}
		return false;
	}

	private async getRecurringOccurrenceDate(): Promise<string | null> {
		let currentValue = moment().format('YYYY-MM-DD');
		while (true) {
			const result = await this.promptForRecurringDate(currentValue);
			if (result === null) {
				return null;
			}
			const trimmed = result.trim();
			if (moment(trimmed, 'YYYY-MM-DD', true).isValid()) {
				return trimmed;
			}
			new Notice('Please enter the date in YYYY-MM-DD format.');
			if (trimmed !== '') { currentValue = trimmed; }
		}
	}

	private async promptForRecurringDate(defaultValue: string): Promise<string | null> {
		return await new Promise((resolve) => {
			const modal = new RecurringOccurrenceModal(this.app, defaultValue, resolve);
			modal.open();
		});
	}

	private applyRecurringOccurrenceDate(fileData: any, occurrenceDate: string): void {
		fileData.helper_selectedOccurrenceDate = occurrenceDate;
		const selectedDate = moment(occurrenceDate, 'YYYY-MM-DD', true);
		if (!selectedDate.isValid()) { return; }

		const originalStart = moment(fileData.apptStartWhole);
		const originalEnd = moment(fileData.apptEndWhole);
		const originalStartLocal = moment(fileData.apptStartWholeLocal);
		const originalEndLocal = moment(fileData.apptEndWholeLocal);

		const durationMs = originalStart.isValid() && originalEnd.isValid()
			? originalEnd.diff(originalStart)
			: (originalStartLocal.isValid() && originalEndLocal.isValid()
				? originalEndLocal.diff(originalStartLocal)
				: null);

		let adjustedStart = originalStart.isValid()
			? originalStart.clone()
			: (originalStartLocal.isValid()
				? originalStartLocal.clone()
				: selectedDate.clone().startOf('day'));
		adjustedStart = adjustedStart
			.year(selectedDate.year())
			.month(selectedDate.month())
			.date(selectedDate.date());
		fileData.apptStartWhole = adjustedStart.format();

		let adjustedEnd: moment.Moment | null = null;
		if (originalEnd.isValid()) {
			adjustedEnd = originalEnd.clone()
				.year(selectedDate.year())
				.month(selectedDate.month())
				.date(selectedDate.date());
		} else if (durationMs !== null) {
			adjustedEnd = adjustedStart.clone().add(durationMs);
		}
		if (adjustedEnd) {
			fileData.apptEndWhole = adjustedEnd.format();
		}

		if (originalStartLocal.isValid()) {
			const adjustedStartLocal = originalStartLocal.clone()
				.year(selectedDate.year())
				.month(selectedDate.month())
				.date(selectedDate.date());
			fileData.apptStartWholeLocal = adjustedStartLocal.format();
		} else {
			fileData.apptStartWholeLocal = adjustedStart.format();
		}

		if (originalEndLocal.isValid()) {
			let adjustedEndLocal = originalEndLocal.clone()
				.year(selectedDate.year())
				.month(selectedDate.month())
				.date(selectedDate.date());
			if (durationMs !== null && originalStartLocal.isValid()) {
				const startLocalAdjusted = originalStartLocal.clone()
					.year(selectedDate.year())
					.month(selectedDate.month())
					.date(selectedDate.date());
				adjustedEndLocal = startLocalAdjusted.add(durationMs);
			} else if (durationMs !== null) {
				adjustedEndLocal = adjustedStart.clone().add(durationMs);
			}
			fileData.apptEndWholeLocal = adjustedEndLocal.format();
		} else if (durationMs !== null && adjustedEnd) {
			fileData.apptEndWholeLocal = adjustedEnd.format();
		}

		if (adjustedStart.isValid()) {
			fileData.helper_selectedOccurrenceDateTimeIso = adjustedStart.toISOString();
		}

		const adjustedStartMoment = moment(fileData.apptStartWhole);
		if (adjustedStartMoment.isValid()) {
			for (const key of Object.keys(fileData)) {
				if (!(typeof fileData[key] === 'string')) { continue; }
				const lowerKey = key.toLowerCase();
				if (lowerKey.startsWith('apptstart') && lowerKey.includes('date')) {
					fileData[key] = adjustedStartMoment.format('YYYY-MM-DD');
				} else if (lowerKey.startsWith('apptstart') && lowerKey.includes('time')) {
					fileData[key] = adjustedStartMoment.format('HH:mm');
				} else if (lowerKey.startsWith('apptstart') && lowerKey.includes('text')) {
					fileData[key] = adjustedStartMoment.format('L LT');
				}
			}
		}

		if (typeof fileData.apptEndText === 'string' && fileData.apptEndWhole) {
			const adjustedEndMoment = moment(fileData.apptEndWhole);
			if (adjustedEndMoment.isValid()) {
				fileData.apptEndText = adjustedEndMoment.format('L LT');
			}
		}
	}

	// Custom funciton to handle a file being dropped onto the ribbon icon.
	async handleDropEvent(dropevt: DragEvent) {
		if (dropevt.dataTransfer == null) {
			throw new ReferenceError('Outlook Meeting Notes cannot handle the DragEvent. The event had a null '
				+ 'dataTransfer property, which should never happen when dispatched by the browser, according '
				+ 'to https://developer.mozilla.org/en-US/docs/Web/API/DragEvent/dataTransfer');
		} else {
			const droppedFiles = dropevt.dataTransfer.files
			if (droppedFiles.length != 1) {
				new Notice('Outlook Meeting Notes can only handle one meeting being dropped onto the ribbon icon');
			}
			else {
				// One file was dropped, hand it over to createMeetingNote
				const droppedFile = droppedFiles[0];

				const fr = new FileReader();
				fr.onload = () => {
					if (fr.result == null) {
						throw new ReferenceError('Outlook Meeting Notes cannot handle the DragEvent. The FileReader had '
							+ 'a null result property, which should not be possible.');
					} else if (!(fr.result instanceof ArrayBuffer)) {
						throw new TypeError('Outlook Meeting Notes cannot handle the DragEvent. The FileReader result '
							+ 'property was not an ArrayBuffer, which should be impossible.');
					} else {
						// As readAsArrayBuffer is being used, below, fr.result will be an ArrayBuffer.
						const msgRdr = new MsgReader(fr.result);
						this.createMeetingNote(msgRdr);
					}
				}
				fr.readAsArrayBuffer(droppedFile)
			}
		}
	}

	private ribbonIconEl: HTMLElement;

	async onload() {
		await this.loadSettings();

		const tooltipMessage = 'Outlook Meeting Notes: Drag and drop a meeting onto this icon from Outlook (or a .msg file) to create a meeting note.';

		// This creates an icon in the left ribbon.
		// Create an icon that does nothing when clicked, as the effect is from 
		this.ribbonIconEl = this.addRibbonIcon('calendar-clock', tooltipMessage, () => { });

		// These respond to something being dragged over the ribbon icon and dropped onto it
		this.ribbonIconEl.addEventListener('dragenter', () => {
			// Style for is-being-dragged-over is defined in the CSS file.
			this.ribbonIconEl.toggleClass('is-being-dragged-over', true);
			// Display a tooltip.
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
			// Remove the tooltip when the user leaves the ribbon icon while still dragging.
			const tooltip = document.getElementsByClassName('tooltip')[0]
			if (tooltip) { tooltip.remove(); }
		});
		this.ribbonIconEl.addEventListener('dragover', (dragevt) => {
			// User is dragging something over the icon

			// Prevent the default behaviour of not allowing the drop event
			dragevt.preventDefault();
			// The dropEffect changes the cursor. Options are 'none', 'move', 'copy', and 'link'.
			if (dragevt.dataTransfer != null) { dragevt.dataTransfer.dropEffect = 'copy'; }
		});
		this.ribbonIconEl.addEventListener('drop', (dropevt) => {
			// User has dropped something on the icon

			// If the user drops something, dragleave doesn't get called.
			this.ribbonIconEl.toggleClass('is-being-dragged-over', false);

			// Prevent the default behaviour because we're handling it.
			dropevt.preventDefault();
			// Call the custom function defined above.
			this.handleDropEvent(dropevt);
		});
		this.ribbonIconEl.addClass('outlook-meeting-notes-icon');

		// This adds a status bar item to the bottom of the app. Does not work on mobile apps.
		// const statusBarItemEl = this.addStatusBarItem();
		// statusBarItemEl.setText('Status Bar Text');

		//TODO: Add commands to create meeting notes.

		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new OutlookMeetingNotesSettingTab(this.app, this));
	}

	onunload() {
		//Don't need to remove event listeners because the icon will be gone.
	}

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
					const rawValue = render(parts[0]).trim().replace(/^"|"$/g, '');
					let formattedMoment = moment(rawValue);
					if (!formattedMoment.isValid()) {
						const ctx = this as Record<string, unknown> & { helper_selectedOccurrenceDateTimeIso?: string; helper_selectedOccurrenceDate?: string; };
						const isoFallback = typeof ctx.helper_selectedOccurrenceDateTimeIso === 'string' ? ctx.helper_selectedOccurrenceDateTimeIso : undefined;
						if (isoFallback) { formattedMoment = moment(isoFallback); }
						if (!formattedMoment.isValid() && typeof ctx.helper_selectedOccurrenceDate === 'string') {
							formattedMoment = moment(ctx.helper_selectedOccurrenceDate, 'YYYY-MM-DD', true);
						}
					}
					if (!formattedMoment.isValid()) { return rawValue; }
					return formattedMoment.format(parts[1]);
				}
			}
		};
		let func: 'firstWord' | 'dateFormat';
		for (func in helperFunctions) {
			hash['helper_' + func] = helperFunctions[func];
		}
		// Add helper functions to all objects in arrays so that the helper 
		// functions work inside mustache sections
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

	//////////////////////
	// renderTemplate() //
	//////////////////////
	// Parse template into YAML and markdown sections to use different escaping for each
	renderTemplate(template: string, hash: any): string {
		// Regex /^---(\r\n?|\n).*?(\r\n?|\n)---($|\r\n?|\n)/s matches '---' at the start of the string,
		// then a platform-independent new-line, followed by a non-greedy match (because of *?) of any
		// character including newlines (because of /s at end), followed by another --- on its own line
		// with either another platform-independent new-line afterwards or the end of the string.
		const templateYAMLMatch = template.match(/^---(\r\n?|\n).*?(\r\n?|\n)---($|\r\n?|\n)/s);
		// If there was a match, then make notesTemplateMD everything after it, otherwise it's the whole thing.
		const templateMD = templateYAMLMatch ? template.substring(templateYAMLMatch[0].length) : template;

		let output = ''

		if (templateYAMLMatch) {
			const sanitizeYamlValue = (value: string): string => value.replace(/[><*]/g, '');
			const mustacheYAMLOptions = {
				escape: (str: string): string => {
					// Escape YAML
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

}

class RecurringOccurrenceModal extends Modal {
	private readonly defaultDate: string;
	private readonly resolvePromise: (value: string | null) => void;
	private input!: TextComponent;
	private hasResolved = false;

	constructor(app: App, defaultDate: string, resolve: (value: string | null) => void) {
		super(app);
		this.defaultDate = defaultDate;
		this.resolvePromise = resolve;
	}

	onOpen(): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Recurring meeting' });
		contentEl.createEl('p', { text: 'This event is recurring. Enter the occurrence date in YYYY-MM-DD format.' });

		const inputWrapper = contentEl.createDiv({ cls: 'outlook-meeting-notes-recurring-date-input' });
		this.input = new TextComponent(inputWrapper);
		this.input.setPlaceholder('YYYY-MM-DD');
		this.input.setValue(this.defaultDate);
		this.input.inputEl.setAttribute('aria-label', 'Recurring meeting date');
		this.input.inputEl.addEventListener('keydown', (evt) => {
			if (evt.key === 'Enter') {
				evt.preventDefault();
				this.submit();
			}
		});

		const buttonWrapper = contentEl.createDiv({ cls: 'outlook-meeting-notes-recurring-date-buttons' });
		new ButtonComponent(buttonWrapper)
			.setButtonText('Cancel')
			.onClick(() => this.cancel());
		new ButtonComponent(buttonWrapper)
			.setButtonText('OK')
			.setCta()
			.onClick(() => this.submit());

		setTimeout(() => {
			this.input.inputEl.focus({ preventScroll: true });
			this.input.inputEl.select();
		}, 0);
	}

	private submit(): void {
		if (this.hasResolved) { return; }
		this.hasResolved = true;
		const value = this.input.getValue().trim();
		this.resolvePromise(value);
		this.close();
	}

	private cancel(): void {
		if (this.hasResolved) { return; }
		this.hasResolved = true;
		this.resolvePromise(null);
		this.close();
	}

	onClose(): void {
		if (!this.hasResolved) {
			this.hasResolved = true;
			this.resolvePromise(null);
		}
		this.contentEl.empty();
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
				link.href = 'https://www.github.com/davidingerslev/outlook-meeting-notes/#outlook-meeting-notes';
				link.target = '_blank';
				link.rel = 'noopener';
				link.textContent = 'documentation';
				df.appendChild(link);
				df.appendChild(document.createTextNode('.'))
				return df;
			})())

	}
}

/*
 * Notes from the sample plugin README.md

## Releasing new releases

- Update your `manifest.json` with your new version number, such as `1.0.1`, and the minimum Obsidian version required for your latest release.
- Update your `versions.json` file with `"new-plugin-version": "minimum-obsidian-version"` so older versions of Obsidian can download an older version of your plugin that's compatible.
- Create new GitHub release using your new version number as the "Tag version". Use the exact version number, don't include a prefix `v`. See here for an example: https://github.com/obsidianmd/obsidian-sample-plugin/releases
- Upload the files `manifest.json`, `main.js`, `styles.css` as binary attachments. Note: The manifest.json file must be in two places, first the root path of your repository and also in the release.
- Publish the release.

> You can simplify the version bump process by running `npm version patch`, `npm version minor` or `npm version major` after updating `minAppVersion` manually in `manifest.json`.
> The command will bump version in `manifest.json` and `package.json`, and add the entry for the new version to `versions.json`

## Adding your plugin to the community plugin list

- Check the [plugin guidelines](https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines).
- Publish an initial version.
- Make sure you have a `README.md` file in the root of your repo.
- Make a pull request at https://github.com/obsidianmd/obsidian-releases to add your plugin.
 */


