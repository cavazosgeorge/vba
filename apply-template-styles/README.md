# Apply Template Styles

A VBA macro that copies header/footer content, syncs header/footer font type, applies the template's body font, forces text color to black, and formats table headers (first row gray + bold) in your target document in one click. Works with OneDrive/SharePoint files.

## What It Does

1. **Creates a timestamped backup** of your document before making any changes
2. **Copies headers and footers** from the template (Primary, First Page, and Even Pages variants)
3. **Syncs header/footer font type** from template to target
4. **Sets text font family** in the target document to the template's `Normal` style font
5. **Sets text color to black** in the target document body text (including notes/comments/text boxes)
6. **Formats all table headers** in the target document: first row gets gray fill (`RGB(191, 191, 191)`) and bold text
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

Headers/footers are matched **by section index**. If your document has more sections than the template, the extra sections keep their existing formatting. If the template has more, the extras are ignored.

## Why Open Both Documents?

This macro avoids file picker paths entirely. OneDrive and SharePoint synced files return URL-encoded paths that Word's VBA APIs often reject. By having both documents already open in Word, we bypass all path-related issues and work directly with the in-memory document objects.

## Notes

- The template can be a `.docx`, `.dotx`, or `.dotm` file — used as the header/footer source
- Works with OneDrive, SharePoint, and local files
- Always test on copies of real documents first
- The backup is saved in the same folder as the original document
- If anything goes wrong mid-run, the error dialog includes the backup path so you can restore
