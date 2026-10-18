# nvim-docx

[![Stargazers](https://img.shields.io/github/stars/DaanHessen/nvim-docx?style=flat-square)](https://github.com/DaanHessen/nvim-docx/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://github.com/DaanHessen/nvim-docx/blob/main/LICENSE)
[![Neovim](https://img.shields.io/badge/Neovim-0.8+-green.svg?style=flat-square)](https://neovim.io)

Open and edit `.docx` files in Neovim as Markdown — **with zero external dependencies**.

`nvim-docx` includes a complete, pure-Lua OOXML (`.docx`) and PKZIP/DEFLATE engine. When you open a Word document, it parses the `.docx` archive directly in Lua and loads it as clean Markdown. When you save (`:w`), it rebuilds the `.docx` while preserving the original document's styles, fonts, headers, footers, and page margins.

---

## Features

- **Zero external dependencies**: Written in 100% pure Lua. No Pandoc, Python, `unzip`, or system packages required — just add the plugin to your Neovim config.
- **No manual conversion**: Run `nvim document.docx` directly. Neovim loads the text as Markdown, and `:w` saves back to `.docx`.
- **Full style & package preservation**: Preserves your original document's `styles.xml`, themes, fonts, headers, footers, custom XML, and page layout (`sectPr`) on save.
- **Rich Markdown & Word support**:
  - Headings (`#` through `######`), localized Word heading styles, and automatic heading bookmarks
  - **Bold**, *italic*, ***bold italic***, ~~strikethrough~~, `inline code`, ^superscript^, and ~subscript~
  - Unordered lists, ordered lists, nested sub-lists, and task lists (`- [x]` / `- [ ]`)
  - Clean Markdown pipe tables (`| Col 1 | Col 2 |`) with column alignment (`:---`, `:---:`, `---:`)
  - Fenced code blocks (` ``` `) and blockquotes (`>`)
  - Hyperlinks (`[text](url)`) and internal section anchors (`[text](#heading)`)
  - Footnotes and endnotes (`[^1]` and `[^1]: ...`)
  - Horizontal rules (`---`) and tracked changes (`accept`, `reject`, `all`)
- **Embedded image support**: Extracts media to a temporary workspace on open and packs existing or newly referenced images (PNG, JPEG, GIF, BMP, WebP, SVG) back into the `.docx` archive with automatic dimension detection.
- **Safe atomic saves**: Compiles to a temporary file first before atomically replacing the target on disk, so an interrupted write never corrupts your document.
- **Built-in commands**: Inspect document statistics (`:DocxInfo`), export `.docx` to Markdown (`:DocxExportMarkdown`), or compile any Markdown buffer directly to `.docx` (`:DocxFromMarkdown`).

---

## Requirements

- **Neovim** >= 0.8.0 (no external tools or binaries needed)

---

## Installation

### [lazy.nvim](https://github.com/folke/lazy.nvim) / LazyVim

```lua
{
  "DaanHessen/nvim-docx",
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

---

## Configuration

`nvim-docx` works out of the box with zero configuration. To customize any settings, pass an options table to `setup` (or `opts` in `lazy.nvim`):

```lua
require("nvim-docx").setup({
  -- Base directory where per-document workspaces are created
  temp_dir = vim.fn.stdpath("cache") .. "/nvim-docx",

  -- Markdown dialect used in Neovim ("markdown" or "gfm")
  markdown_format = "markdown",

  -- Delete workspace cache and extracted media when the buffer is closed
  auto_cleanup = true,

  -- Retain original Word styles, headers, footers, themes, and page setup on save
  preserve_styles = true,

  -- Line wrap mode ("none", "auto", "preserve")
  wrap = "none",

  -- Track-changes policy ("accept", "reject", "all", or nil for default "accept")
  track_changes = nil,

  -- Include YAML frontmatter for DOCX core document properties (title, author, subject, keywords)
  frontmatter = false,

  -- Print verbose notifications for debugging
  debug = false,
})
```

---

## Commands

| Command | Description |
| :--- | :--- |
| `:DocxInfo` | Display word count, character count, block statistics, and metadata for the current `.docx` buffer. |
| `:DocxExportMarkdown [path]` | Export the current `.docx` buffer (and any extracted `media/` files) to a `.md` file. |
| `:DocxFromMarkdown [path]` | Compile the current Markdown buffer directly into a `.docx` file. |

---

## How it works

1. A `BufReadCmd` autocommand intercepts `.docx` files before Neovim reads them as binary.
2. A temporary workspace folder is created in `stdpath("cache")/nvim-docx/<name>-<id>/` to hold extracted images.
3. The pure-Lua ZIP (`nvim-docx.zip`) and OOXML (`nvim-docx.ooxml`) engine decompresses the archive, extracts `word/media/*` into the workspace, converts `word/document.xml` (along with styles, numbering, relationships, and footnotes) to Markdown, and sets the buffer to `filetype=markdown` and `buftype=acwrite`.
4. When you save (`:w`), `BufWriteCmd` parses the Markdown buffer into an AST, rebuilds the OOXML package (cloning styles, headers, footers, themes, and section layout from the original `.docx` when `preserve_styles = true`), compresses it into a temporary `.docx`, and atomically replaces the target file on disk.
5. Closing or wiping the buffer deletes the workspace folder (unless `auto_cleanup = false`).

---

## Health check

Run `:checkhealth nvim-docx` in Neovim to verify your environment:

```
nvim-docx: require("nvim-docx.health").check()

- OK Neovim version: 0.12.6
- OK Native Lua OOXML engine active (LuaJIT 2.1, zero external dependencies)
- OK Workspace cache directory writable: /home/user/.cache/nvim/nvim-docx
```

---

## Limitations

`nvim-docx` supports headings, inline formatting, lists, task lists, tables, code blocks, blockquotes, links, footnotes, and images. However:
- Proprietary Word structures (such as SmartArt, embedded VBA macros, or complex vector drawing shapes) have no Markdown equivalents and are omitted from the Markdown view.

---

## License

Distributed under the [MIT License](LICENSE).
