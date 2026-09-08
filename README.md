# nvim-docx

[![Stargazers](https://img.shields.io/github/stars/DaanHessen/nvim-docx?style=flat-square)](https://github.com/DaanHessen/nvim-docx/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://github.com/DaanHessen/nvim-docx/blob/main/LICENSE)
[![Neovim](https://img.shields.io/badge/Neovim-0.8+-green.svg?style=flat-square)](https://neovim.io)

Edit Microsoft Word (`.docx`) documents directly in Neovim as Markdown, converting seamlessly on open and save.

Powered by [Pandoc](https://pandoc.org) and native Neovim autocommands (`BufReadCmd` / `BufWriteCmd`).

---

## Features

* **Transparent editing**: Open any `.docx` file in Neovim (`nvim document.docx`) — it renders immediately as editable Markdown.
* **Non-destructive saving**: Saving (`:w`) compiles Markdown back to `.docx` atomically, leaving the original file untouched until conversion succeeds.
* **Style preservation**: Uses the original document as a Pandoc `--reference-doc`, preserving fonts, headings, and margins.
* **Media handling**: Embedded figures and images are extracted to an isolated workspace cache and bundled back into the `.docx` archive upon save.
* **Safe execution**: Runs via `vim.system` without spawning subshells or locking the editor UI.
* **Diagnostics**: Built-in `:checkhealth nvim-docx` to verify environment dependencies.

---

## Requirements

* **Neovim** >= 0.8.0 (0.10+ recommended)
* **[Pandoc](https://pandoc.org/)** available in your `PATH`

---

## Installation

### [lazy.nvim](https://github.com/folke/lazy.nvim)

```lua
{
  "DaanHessen/nvim-docx",
  ft = "docx",
  opts = {
    -- optional configuration
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

`nvim-docx` works out of the box with sensible defaults without calling `.setup()`. To customize options:

```lua
require("nvim-docx").setup({
  -- Path or command name for the pandoc binary
  pandoc_path = "pandoc",

  -- Base directory where isolated per-document workspaces are created
  temp_dir = vim.fn.stdpath("cache") .. "/nvim-docx",

  -- Markdown dialect used in Neovim ("markdown", "gfm", etc.)
  markdown_format = "markdown",

  -- Wipe workspace cache and temporary media when the buffer is closed
  auto_cleanup = true,

  -- Retain original Word formatting and typography via Pandoc reference-doc
  preserve_styles = true,

  -- Markdown line wrap mode ("none", "auto", "preserve")
  wrap = "none",

  -- Track-changes policy ("accept", "reject", "all", or nil to use pandoc default)
  track_changes = nil,

  -- Output verbose debugging notifications
  debug = false,
})
```

---

## How It Works

1. **Interception**: When a `.docx` buffer opens, a `BufReadCmd` autocommand intercepts the event before Neovim reads the binary ZIP container.
2. **Workspace**: An isolated temporary directory is created inside `temp_dir` (`~/.cache/nvim/nvim-docx/<name>-<id>/`).
3. **Extraction**: Pandoc converts the document into Markdown and extracts any embedded images into the workspace.
4. **Virtual buffer**: Buffer content is populated with Markdown text, and `buftype` is set to `acwrite`.
5. **Atomic write**: On `:w`, a `BufWriteCmd` autocommand serializes the buffer and invokes Pandoc to build a temporary `.docx` using the original file as `--reference-doc`. Once successful, it atomically replaces the target file.
6. **Cleanup**: When `auto_cleanup = true`, closing or wiping the buffer (`BufWipeout`) removes the temporary workspace.

---

## Health Check

Run `:checkhealth nvim-docx` in Neovim to verify your setup:

```
nvim-docx: require("nvim-docx.health").check()

- OK Neovim version: 0.12.5
- OK Pandoc executable found at 'pandoc' (3.10.2)
- OK Workspace cache directory writable: /home/user/.cache/nvim/nvim-docx
```

---

## Limitations

* Complex proprietary Word structures (such as SmartArt, embedded macros, or ActiveX controls) cannot be represented in Markdown and will be omitted by Pandoc.
* Conversion fidelity depends on Pandoc's OOXML and Markdown capabilities.

---

## License

Distributed under the [MIT License](LICENSE).
