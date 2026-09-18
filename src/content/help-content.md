# Welcome to StyleMash

StyleMash cleans up messy Word documents by finding every distinct text style already in a file - named or not - and letting you fold groups of them into a small set of clean, named styles you control.

Everything happens in your browser. Your document is never uploaded or sent anywhere; closing the tab is the only "delete" you need.

## Getting started

Drag a .docx or .dotx file onto the window, or click to browse for one. StyleMash reads it and switches straight to the workspace - nothing else to configure first.

Only .docx and .dotx are supported (the modern Office Open XML format). An old .doc file from before Word 2007 needs to be re-saved as .docx first.

## Current styles: what's actually in your document

The left-hand "Current styles" panel lists every distinct text appearance StyleMash found in the document - most common first - however it got its look, whether that's a named Word style or formatting applied by hand with no style behind it. Each entry shows a live preview, its resolved font, size, colour, bold, italic, underline and strikethrough, and where that look came from.

Tick the checkbox on any entry, or click the row, to select it. Select as many entries as you like before merging them together.

## New Styles: where merged text ends up

The right-hand "New Styles" panel is the list of named styles you can merge into. It starts empty - populate it however suits you:

- Click "+ New Style" to define one from scratch.
- Click "+ Defaults" to add a bundled set of common starter styles (headings, bulleted and numbered lists, Normal, and a few text styles) in one go.
- Click "Attach custom Word styles" to import every named style from a second, already-clean Word file, so you can reuse styles you have already approved elsewhere.

Click "Edit" on any New Styles entry later to change its look.

## Merging styles

There are two ways to merge selected entries into a style:

1. Select one or more entries on the left, then click a style's row on the right to mark it as the target - its background turns amber to show it is picked. Click "Merge N selected here" on that row, or click "Mash it" on the left, which does the same thing once a target is already picked. This merges immediately, with no extra dialog.
2. Select one or more entries on the left and click "Mash it" without picking a target first. A dialog opens where you can either create a brand-new style on the spot or pick an existing one from New Styles, then confirm.

Either way, every selected entry's text is re-styled to match the target style, which now covers that formatting wherever it appears in the document.

## Choosing which bundled defaults you get

Not every default style is useful for every document. Click "Customise your own style file" in the top header to open a checklist of every bundled default, grouped into "Body text styles", "Heading Styles" and "List styles". Untick anything you do not want, and the next "+ Defaults" click only adds what is still ticked. The same panel has a link to download the underlying reference Word file, if you would rather open it in Word and build your own variant.

## Using a second Word file as your style source

If you already have a clean, approved Word document with the styles you want everything to match, click "Attach custom Word styles" in the New Styles panel footer. Every named style it defines is imported as an extra merge target, tagged "from Document B" so you can tell them apart from styles you built by hand. Click "Remove Document B" to detach it again - any of its styles you never actually merged into are removed along with it, so nothing unused is left behind.

## Undo, clearing, and starting over

"Undo", next to "Mash it", reverts the most recent merge, one step at a time. "Clear list" wipes every New Styles entry you have created this session. "Mash a different file" takes you back to the upload screen to start over with a new document.

## Saving your result

Click "Save your file" at any time to download the cleaned document. It keeps the original file format and filename, with "-RIPPED" appended - for example, Report.docx becomes Report-RIPPED.docx. Saving does not end your session; keep merging and save again whenever you like.

## What StyleMash does not do yet

Headers, footers, and footnotes or endnotes are not scanned - only the main body text. Page layout, tables and images are not reproduced in the Document Preview; it is a formatting preview, not a full page-accurate rendering. There is no mobile-friendly version planned.
