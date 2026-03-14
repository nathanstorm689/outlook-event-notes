# Outlook Event Notes

An [Obsidian](https://obsidian.md) plugin that creates notes from Microsoft Outlook meetings, appointments, and recurring events by dragging and dropping `.msg` files onto a ribbon icon.

> **Based on [Outlook Meeting Notes](https://github.com/davidingerslev/outlook-meeting-notes) by [David Ingerslev](https://github.com/davidingerslev).**
> This fork adds support for recurring events, opening existing notes, and other improvements.

---

## Features

- **Drag & drop** a meeting or appointment from Outlook Classic onto the ribbon icon → a note is instantly created (or opened if it already exists)
- **Recurring events** — detects recurring `.msg` files and prompts you to confirm the occurrence date before creating the note
- **Fully customisable** filename pattern and note template using [Mustache](https://mustache.github.io/mustache.5.html) syntax
- **No Microsoft 365 / Graph API dependency** — works entirely from `.msg` files
- **YAML frontmatter sanitisation** — strips characters that would break Obsidian properties

## Installation

### Community Plugins (coming soon)
Search for **Outlook Event Notes** in Obsidian → Settings → Community Plugins.

### Manual installation
1. Download `main.js`, `manifest.json`, and `styles.css` from the [latest release](https://github.com/nathanstorm689/outlook-event-notes/releases/latest)
2. Copy them into your vault at `.obsidian/plugins/outlook-event-notes/`
3. Enable the plugin in Obsidian → Settings → Community Plugins

---

## Usage

Drag and drop a meeting or appointment from the **Outlook Classic** desktop calendar onto the plugin ribbon icon. The plugin will:

1. Parse the `.msg` file
2. For recurring events, ask you to confirm the occurrence date (defaults to today)
3. Create a new note — or open the existing one if a note for that event already exists

You can also save a `.msg` file from Outlook (e.g. a meeting invitation received by email) and drag-drop the saved file onto the icon.

### Recurring events
If the dropped `.msg` represents part of a recurring series, the plugin shows a modal asking for the occurrence date in `YYYY-MM-DD` format. Press Enter to accept the default (today), adjust the date if needed, or cancel to abort note creation.

---

## Settings

### Folder location
The folder where new notes are created. Created automatically if it does not exist (supports subfolders like `Meetings/2026`).

### Filename pattern
Uses Mustache syntax. Default:
```
{{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD_HH-mm_ss{{/helper_dateFormat}} {{subject}}
```
Produces filenames like `2026-03-11_19-00_00 Discuss documentation`.

You can use `/` to create subfolders:
```
{{#helper_dateFormat}}{{apptStartWhole}}|YYYY/MM/YYYY-MM-DD_HH-mm_ss{{/helper_dateFormat}} {{subject}}
```
produces `2026/03/2026-03-11_19-00_00 Discuss documentation`

### Invalid character substitute
Characters that are invalid in filenames (`/ * " \ < > : | ?`) are replaced with this value. Blank = remove them.

### Template
Fully customisable Mustache template. All `.msg` [fields](https://hiraokahypertools.github.io/msgreader/typedoc/interfaces/MsgReader.FieldsData.html) are available, plus the helper fields and functions below.

YAML values are sanitised automatically to strip characters such as `<`, `>` and `*`, keeping the generated frontmatter valid even when invite data contains those symbols.

#### Default template
```
---
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
```

---

## Template helpers

### Helper fields

| Field | Description |
|---|---|
| `{{helper_currentDT}}` | Date/time the file was dropped, in ISO format (e.g. `2026-03-11T19:00:00-05:00`) |

Use `helper_dateFormat` to reformat it:
```
{{#helper_dateFormat}}{{helper_currentDT}}|YYYY-MM-DD_HH-mm_ss{{/helper_dateFormat}}
```

### Helper functions

#### `helper_dateFormat`
Formats a date using [moment.js](https://momentjs.com/). Separate the field and format with `|`.

```
{{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD_HH-mm_ss{{/helper_dateFormat}}
```
produces `2026-03-11_19-00_00`

```
{{#helper_dateFormat}}{{apptStartWhole}}|L LT{{/helper_dateFormat}}
```
Uses Obsidian's display language locale (e.g. `11/03/2026 19:00` in English GB, `03/11/2026 7:00 PM` in English US).

#### `helper_firstWord`
Returns only the first word of a field — useful for recording first names only:
```
meeting-recipients:
{{#recipients}}
  - {{#helper_firstWord}}{{name}}{{/helper_firstWord}}
{{/recipients}}
```

---

## Credits

- Original plugin: [Outlook Meeting Notes](https://github.com/davidingerslev/outlook-meeting-notes) by [David Ingerslev](https://github.com/davidingerslev)
- [msgreader](https://github.com/HiraokaHyperTools/msgreader) — `.msg` file parsing
- [mustache.js](https://github.com/janl/mustache.js) — template rendering
- [mustache-validator](https://github.com/eliasm307/mustache-validator) — template validation

---

## License

ISC — see [LICENSE](LICENSE).
