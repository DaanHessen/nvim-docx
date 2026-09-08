# nvim-docx

[![Stargazers](https://img.shields.io/github/stars/DaanHessen/nvim-docx?style=flat-square)](https://github.com/DaanHessen/nvim-docx/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://github.com/DaanHessen/nvim-docx/blob/main/LICENSE)
[![Neovim](https://img.shields.io/badge/Neovim-0.8+-green.svg?style=flat-square)](https://neovim.io)

Open and edit `.docx` files in Neovim as Markdown.

When you open a Word document, `nvim-docx` converts it to Markdown in the background using Pandoc. When you save (`:w`), it rebuilds the `.docx` using the original file as a reference template to keep fonts, headings, and margins intact.

---

## Features

- **No manual conversion**: Run `nvim document.docx` directly. Neovim loads the text as Markdown, and `:w` saves back to `.docx`.
- **Style preservation**: Passes your original file to Pandoc as a `--reference-doc` so fonts, heading colors, and margins do not reset to Word defaults.
- **Image support**: Embedded figures and images extract to a temporary workspace and pack back into the `.docx` archive on save.
- **Safe saves**: Compiles to a temporary file first before replacing the original on disk, so a failed conversion will not corrupt your document.
- **No UI lockup**: Runs Pandoc via `vim.system` without spawning subshells or freezing the editor.
- **Health check**: Run `:checkhealth nvim-docx` to verify that Pandoc and cache directories are set up correctly.

---

## Requirements

- **Neovim** >= 0.8.0 (0.10+ recommended)
- **[Pandoc](https://pandoc.org/)** installed and available in your `PATH`

---

## Installation

### [lazy.nvim](https://github.com/folke/lazy.nvim)

```lua
{
  "DaanHessen/nvim-docx",
  ft = "docx",
  opts = {
    -- optional custom settings
  },
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

`nvim-docx` works out of the box with defaults. To change any settings, pass an options table to `setup`:

```lua
require("nvim-docx").setup({
  -- Path or command name for the pandoc binary
  pandoc_path = "pandoc",

  -- Base directory where per-document workspaces are created
  temp_dir = vim.fn.stdpath("cache") .. "/nvim-docx",

  -- Markdown dialect used in Neovim ("markdown", "gfm", etc.)
  markdown_format = "markdown",

  -- Delete workspace cache and extracted media when the buffer is closed
  auto_cleanup = true,

  -- Retain original Word formatting and typography via Pandoc reference-doc
  preserve_styles = true,

  -- Line wrap mode passed to pandoc ("none", "auto", "preserve")
  wrap = "none",

  -- Track-changes policy ("accept", "reject", "all", or nil to use pandoc default)
  track_changes = nil,

  -- Print verbose notifications for debugging
  debug = false,
})
```

---

## How it works

1. A `BufReadCmd` autocommand intercepts `.docx` files before Neovim reads them as binary.
2. A temporary workspace folder is created in `stdpath("cache")/nvim-docx/<name>-<id>/` to hold extracted images and converted Markdown.
3. Pandoc extracts media into this folder, converts the text, and sets the buffer to `filetype=markdown` and `buftype=acwrite`.
4. When you save (`:w`), `BufWriteCmd` exports the buffer to Markdown, runs Pandoc with `--reference-doc` pointing to your original document, and writes a temporary `.docx`. Once Pandoc succeeds, it atomically replaces the target file on disk.
5. Closing or wiping the buffer deletes the workspace folder (unless `auto_cleanup = false`).

---

## Health check

Run `:checkhealth nvim-docx` in Neovim to verify your environment:

```
nvim-docx: require("nvim-docx.health").check()

- OK Neovim version: 0.12.5
- OK Pandoc executable found at 'pandoc' (3.10.2)
- OK Workspace cache directory writable: /home/user/.cache/nvim/nvim-docx
```

---

## Limitations

Pandoc handles standard text, headings, lists, tables, footnotes, and images well. However:
- Proprietary Word structures (such as SmartArt, embedded macros, or drawing shapes) have no Markdown equivalents and will be omitted by Pandoc.
- Tracked changes and comments cannot be edited through Markdown.

---

## License

Distributed under the [MIT License](LICENSE).
