# Apply Template Styles

A VBA macro that enforces consistent formatting across Word documents by attaching a template and propagating its styles, headers/footers, and page layout in one click.

## What It Does

1. **Creates a timestamped backup** of your document before making any changes (e.g., `MyDoc_backup_2026-02-24_143022.docx`)
2. **Attaches a `.dotx`/`.dotm` template** to the active document
3. **Updates all styles** — paragraph, character, list, and table styles are overwritten to match the template's definitions
4. **Disables auto-update on open** so styles won't silently change if the document is opened later with a different Normal.dotm
5. **Copies headers and footers** from the template, including Primary, First Page, and Even Pages variants, with Link to Previous preserved
6. **Copies page setup** — margins, orientation, paper size, gutter, header/footer distance, section start type, vertical alignment, and mirror margins
7. **Rebuilds the Table of Contents** if one exists
8. **Updates all fields** (page numbers, cross-references, etc.)
9. **Shows a summary dialog** reporting what was applied

## Installation

1. Open Word
2. Press `Alt+F11` to open the VBA editor
3. Insert → Module
4. Paste the contents of `ApplyTemplateStyles.bas`
5. Close the VBA editor

Alternatively, save it as a `.dotm` add-in in your Word STARTUP folder for persistent access.

## Usage

1. Open your document in Word
2. Press `Alt+F8` → select `ApplyTemplateStyles` → Run
3. Pick your template file when prompted (or hardcode the path — see Configuration)
4. Review the summary dialog

## Configuration

Two constants at the top of the script:

| Constant | Default | Purpose |
|---|---|---|
| `DEFAULT_TEMPLATE_PATH` | `""` (empty) | Set to your template's full path to skip the file picker. Leave empty to be prompted each run. |
| `DISABLE_AUTO_UPDATE` | `True` | Prevents Word from auto-updating styles when the document is opened later. |

## How It Handles Sections

Headers/footers and page setup are matched **by section index**. If your document has more sections than the template, the extra sections keep their existing formatting. If the template has more, the extras are ignored.

## Notes

- Always test on copies of real documents first
- The backup is saved in the same folder as the original document
- If anything goes wrong mid-run, the error dialog includes the backup path so you can restore
