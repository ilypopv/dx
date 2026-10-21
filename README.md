# dx

View **PDF and Word documents** in VS Code. Light on size, and it keeps your documents' formatting intact.

[![Install from the VS Code Marketplace](https://img.shields.io/badge/VS%20Code-Install%20from%20Marketplace-1f9d6b?logo=visualstudiocode&logoColor=white)](https://marketplace.visualstudio.com/items?itemName=ilypopv.dx)

## Features

- Opens `.pdf`, `.docx` and `.doc` as pages, in an extension of under 3 MB.
- PDF: select and copy text, follow links, fill in forms, rotate pages.
- Annotate PDFs (highlight, text, drawing); annotations are saved into the file.
- Edit the text of `.docx` documents; only the edited text changes, styles, images and tables are preserved.
- Headings, lists, tables, images, links and footnotes in DOCX; text, formatting and tables in DOC.
- Find, page and word counts, zoom. The position in every file is remembered.

## Install

Search for **dx** in the Extensions view, or run `ext install ilypopv.dx`.

## Use

- **PDF, DOCX and DOC files** open in dx automatically.
- **Edit a DOCX**: click **Edit** in the toolbar, type, then save. **Annotate a PDF**: pick a tool in the toolbar, then save.

| Action                | Shortcut                 |
| --------------------- | ------------------------ |
| Find                  | `Ctrl/Cmd+F`             |
| Next / previous match | `F3` / `Shift+F3`        |
| Zoom                  | `Ctrl/Cmd` + mouse wheel |
| Leave edit mode       | `Escape`                 |
| Save                  | `Ctrl/Cmd+S`             |
| Open a link in Edit   | `Ctrl/Cmd` + click       |

## Limitations

- DOCX editing changes text within paragraphs only: no new paragraphs, no formatting changes. Headers, footers and comments are not shown.
- DOCX pages follow the page breaks saved by Word, so a page may be longer than on paper.
- DOC files are read-only; Word 6/95 and encrypted documents are not supported.
- Back up important files before editing.

## Settings

### Configuring VS Code

You can make dx the default editor for your document files by setting `workbench.editorAssociations` in `settings.json`:

```json
{
    "workbench.editorAssociations": {
        "*.pdf": "dx-pdf-editor",
        "*.docx": "dx-docx-editor",
        "*.doc": "dx-doc-editor"
    }
}
```

Without it, these files already open in dx; the setting matters when another extension also claims them. To go back to the built-in editor for a type, set its value to `"default"`. To switch for a single file, use **Open With...**.

### Configuring dx

dx has no settings of its own yet.

## Development

```bash
npm install
npm test
npm run package   # builds dx.vsix
```

Press `F5` to debug.

<div align="center">
  <a target="_blank" href="https://ilypopv.github.io">
    <img src="docs/images/made-by-ilypopv.png" alt="Made by ilypopv" width="192">
  </a>
</div>
