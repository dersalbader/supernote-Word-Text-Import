# Word Text Import — Supernote Plugin

Beta test build for Supernote Manta running Chauvet 3.29.44 beta.

Current version: **0.1.9 beta**. The plugin window, instructions and controls now
use solid opaque backgrounds for clearer reading over handwritten or printed pages.

## What it does

1. Selects a `.docx` Word document with Supernote's file picker.
2. Extracts its editable text.
3. Inserts the text directly as an editable text box on the current note page.
4. Copies the text to the clipboard as a fallback if direct insertion is refused.

Paragraphs, line breaks, tabs and table-cell text are retained. Images, page layout,
text styling, headers, footers, footnotes and complex Word formatting are not imported.

## Safety fallback

The plugin uses Supernote's text-box API rather than editing the proprietary `.note`
file itself. If the beta firmware refuses direct insertion, the extracted text is copied
to the clipboard and can still be pasted into a text box.

## Installation and use

Install `WordTextImport-v0.1.9-beta.snplg` through the Supernote plugin installation system.
Open the plugin, select **Choose Word document**, choose a `.docx`, and wait for the
confirmation. The text is inserted into an editable text box on the current page.

Test with a copy of an important note first. This is beta software.
