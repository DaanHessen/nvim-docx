# nvim-docx

[![Stargazers](https://img.shields.io/github/stars/DaanHessen/nvim-docx?style=flat-square)](https://github.com/DaanHessen/nvim-docx/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://github.com/DaanHessen/nvim-docx/blob/main/LICENSE)
[![Neovim](https://img.shields.io/badge/Neovim-0.8+-green.svg?style=flat-square)](https://neovim.io)

Open, edit and create Word documents in Neovim as Markdown, with no external dependencies.

`nvim-docx` contains its own ZIP/DEFLATE, WordprocessingML and legacy Word readers, written in Lua. Run `nvim report.docx`, edit the Markdown you see, and `:w` writes a real Word document back. Everything you did not touch is written back exactly as it was, so fonts, sizes, colors, spacing, headers, footers and page layout survive editing.

---

## Features

- **No dependencies**: no Pandoc, Python, `unzip` or LibreOffice. Install the plugin and open a file.
- **Every Word generation**:
  - `.docx`, `.docm`, `.dotx`, `.dotm` (Word 2007 and later): read and write.
  - `.doc`, `.dot` (Word 97–2003, Word 6.0/95): read with formatting, headings, lists, tables, hyperlinks, footnotes, text boxes and pictures. `:w` saves a `.docx` next to the original and leaves the `.doc` untouched.
  - Word for Windows 1.x/2.0 and Windows Write (`.wri`): text.
- **Lossless editing**: paragraphs, tables and images whose Markdown you did not change are written back byte for byte. A paragraph you edit keeps its original paragraph and font formatting, a new paragraph takes the look of its neighbours, and a new list item continues the list it was added to.
- **Fonts, sizes and colors in Markdown** (see [Formatting](#formatting)): `[text]{font="Georgia" size=14 color=#C00000}` and whole sections with `::: {font="Garamond" size=12 align=justify}`.
- **Markdown coverage**: headings, **bold**, *italic*, <u>underline</u>, ~~strikethrough~~, `code`, ^superscript^, ~subscript~, links, images, bullet/numbered/task lists, pipe tables with alignment, block quotes, code blocks, footnotes, math (`$...$`, `$$...$$`), horizontal rules and page breaks.
- **Live preview**: `:DocxPreview` opens the document in your browser and updates as you type.
- **Fast**: a 15 MB document opens in about 0.1 s. Large images and embedded objects are only decompressed when needed and are written back without recompressing.
- **Safe saves**: the document is written to a temporary file first and then moved over the original, so an interrupted write never leaves a broken file.

---

## Requirements

- **Neovim** 0.8 or newer (tested on 0.8.3, 0.9.5 and 0.12). Nothing else.

---

## Installation

### [lazy.nvim](https://github.com/folke/lazy.nvim) / LazyVim

```lua
{
  "DaanHessen/nvim-docx",
  lazy = false, -- must be loaded before you open a Word file
  opts = {},
}
```

### [pckr.nvim](https://github.com/lewis6991/pckr.nvim) / [packer.nvim](https://github.com/wbthomason/packer.nvim)

```lua
use {
  "DaanHessen/nvim-docx",
  config = function()
    require("nvim-docx").setup()
  end,
}
```

### Upgrading from the Pandoc-based version

Nothing to change. Existing configurations keep working: `pandoc_path` is accepted and ignored (`:checkhealth nvim-docx` mentions that it can be removed), commands and options are the same, and Pandoc no longer needs to be installed.

---

## Configuration

`nvim-docx` works without configuration. To change settings, pass a table to `setup` (or `opts` in lazy.nvim):

```lua
require("nvim-docx").setup({
  -- Base directory for per-document workspaces (extracted images)
  temp_dir = vim.fn.stdpath("cache") .. "/nvim-docx",

  -- Markdown dialect ("markdown" or "gfm"; gfm drops image size attributes)
  markdown_format = "markdown",

  -- Delete a document's workspace when its buffer is closed
  auto_cleanup = true,

  -- Keep the original styles, headers, footers, themes and page setup on save
  preserve_styles = true,

  -- Line wrapping of paragraphs in the Markdown view ("none" or "auto")
  wrap = "none",

  -- Tracked changes: "accept" (default), "reject" or "all"
  track_changes = nil,

  -- Show document properties (title, author, ...) as YAML front matter
  frontmatter = false,

  -- Verbose notifications
  debug = false,
})
```

---

## Formatting

Plain Markdown covers the common cases. Things Markdown has no syntax for use [Pandoc's attribute syntax](https://pandoc.org/MANUAL.html#extension-bracketed_spans), so the files stay readable by other tools.

| What | Markdown |
| :--- | :--- |
| Font, size (pt), color, highlight for some text | `[text]{font="Georgia" size=14 color=#C00000 highlight=yellow}` |
| The same for a whole section | `::: {font="Garamond" size=12 color=#333333}` … `:::` |
| Alignment | `::: {align=center}` (`left`, `center`, `right`, `justify`) |
| A Word paragraph style by name | `::: {style="Quote"}`, `::: Title`, `::: {style="Subtitle"}` |
| Page break | `\newpage` on its own line |
| Underline | `<u>text</u>` or `[text]{.underline}` |

Colors are hex (`#C00000`, `C00000`) or a name (`red`, `blue`, `green`, `orange`, `purple`, `gray`, …). Highlights are Word's highlight colors (`yellow`, `green`, `cyan`, `magenta`, …) or any hex color. In a section, `size` applies to body text; headings keep their own size.

When an existing document is opened, a run whose font, size or color differs from the rest of its paragraph is shown with these attributes (for example a single word made larger in Word), so you can see and change it. The paragraph's own font and size are not shown; they are kept automatically.

Example:

```markdown
::: Title

Quarterly report

:::

# Summary

Revenue grew [11 %]{color=green size=14} this quarter.

::: {font="Georgia" size=13 align=justify}

This section is set in Georgia 13 pt, justified.

:::

\newpage
```

---

## Commands

| Command | Description |
| :--- | :--- |
| `:DocxPreview` | Open a live preview of the document in your browser. It updates as you type. |
| `:DocxPreviewStop` | Stop the live preview. |
| `:DocxInfo` | Word count, character count, block statistics and metadata. |
| `:DocxExportMarkdown [path]` | Write the buffer (and its images) to a `.md` file. |
| `:DocxFromMarkdown [path]` | Compile the current Markdown buffer into a `.docx`. |

`:read file.docx` inserts a Word document as Markdown below the cursor, and `:'<,'>write part.docx` writes the selected lines as a new document.

The preview is an HTML rendering of the Markdown in a page-shaped view, served from a small server on `127.0.0.1` that Neovim runs while the preview is open. It shows fonts, sizes, colors, alignment, tables, lists, images and footnotes. Page breaks and exact line wrapping are Word's job and are only approximated.

---

## How it works

1. A `BufReadCmd` autocommand reads Word files before Neovim would load them as binary data (Neovim's bundled zip plugin is kept away from Word files).
2. The document is parsed in Lua and rendered to Markdown. For every block, the plugin remembers the Word XML it came from.
3. On `:w`, the Markdown is parsed again and compared line by line with what was loaded. Unchanged blocks are written back from the original XML. Changed blocks are generated and inherit the formatting of the block they replace. Styles, numbering, headers, footers, sections, comments and embedded objects are carried over.
4. The package is zipped to a temporary file and moved over the target.

---

## Health check

```
:checkhealth nvim-docx

- OK Neovim version: 0.12.6
- OK Native Lua OOXML engine active (LuaJIT 2.1, zero external dependencies)
- OK Workspace cache directory writable: /home/user/.cache/nvim/nvim-docx
```

---

## Limitations

- Headers and footers are kept but not shown or editable in the Markdown view.
- Text boxes, SmartArt and shapes are shown as their text or image where possible; their layout is kept when the surrounding paragraph is unchanged.
- Equations are converted to and from LaTeX-style text; complex equations may come back simplified after editing them.
- `::: {...}` sections are not reconstructed when a document is reopened. The formatting is in the document, but the section markers are not shown again.
- Password-protected documents cannot be opened.
- `.doc` files are saved as `.docx`; the binary `.doc` format is read-only.
- Word 6.0/95 files are read with basic formatting (bold, italic, underline, tables); Word 1.x/2.0 and Write files as plain text.

---

## License

Distributed under the [MIT License](LICENSE).
