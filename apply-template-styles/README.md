# Apply Template Styles

A VBA macro that copies all styles, headers/footers, and page layout from a template document into your target document in one click. Works with OneDrive/SharePoint files.

## What It Does

1. **Creates a timestamped backup** of your document before making any changes
2. **Copies all styles** from the template using `OrganizerCopy` — paragraph, character, list, and table styles
3. **Copies headers and footers** — Primary, First Page, and Even Pages variants, with Link to Previous preserved
4. **Copies page setup** — margins, orientation, paper size, gutter, header/footer distance, section start type, vertical alignment, and mirror margins
5. **Disables auto-update on open** so styles won't silently change later
6. **Rebuilds the Table of Contents** if one exists
7. **Updates all fields** (page numbers, cross-references, etc.)
8. **Shows a summary dialog** reporting what was applied

## Installation

1. Open Word → press `Alt+F11` to open the VBA editor
2. In the left panel, expand **Normal** → right-click **Modules** → Insert → Module
3. Paste the contents of `ApplyTemplateStyles.bas`
4. Close the VBA editor

Storing the macro in Normal.dotm makes it available from any document without needing `.docm` files.

## Usage

1. **Open both documents** in Word — your target document AND your template document
2. **Click into your target document** (the one you want to format)
3. Press `Alt+F8` → select `ApplyTemplateStyles` → Run
4. Select your template from the list of open documents
5. Review the summary dialog

## Configuration

| Constant | Default | Purpose |
|---|---|---|
| `DISABLE_AUTO_UPDATE` | `True` | Prevents Word from auto-updating styles when the document is opened later. |

## How It Handles Sections

Headers/footers and page setup are matched **by section index**. If your document has more sections than the template, the extra sections keep their existing formatting. If the template has more, the extras are ignored.

## Why Open Both Documents?

This macro avoids file picker paths entirely. OneDrive and SharePoint synced files return URL-encoded paths that Word's VBA APIs often reject. By having both documents already open in Word, we bypass all path-related issues and work directly with the in-memory document objects.

## Notes

- The template can be a `.docx`, `.dotx`, or `.dotm` file — all work as a style source
- Works with OneDrive, SharePoint, and local files
- Always test on copies of real documents first
- The backup is saved in the same folder as the original document
- If anything goes wrong mid-run, the error dialog includes the backup path so you can restore
