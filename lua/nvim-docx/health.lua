local M = {}

function M.check()
  vim.health.start("nvim-docx")

  -- Check Neovim version
  if vim.fn.has("nvim-0.8") == 1 then
    vim.health.ok(string.format("Neovim version: %s", vim.version().string or ">= 0.8"))
  else
    vim.health.error("Neovim 0.8+ is required")
  end

  -- Check Pandoc binary
  local config = require("nvim-docx.config")
  local converter = require("nvim-docx.converter")
  local pandoc_bin = config.options.pandoc_path or "pandoc"
  local available, version_str = converter.check_pandoc_available(true)

  if available then
    vim.health.ok(string.format("Pandoc executable found at '%s' (%s)", pandoc_bin, version_str))
  else
    vim.health.error(
      string.format("Pandoc executable '%s' not found in PATH", pandoc_bin),
      {
        "Install pandoc using your system package manager (e.g. `sudo pacman -S pandoc`, `sudo apt install pandoc`, `brew install pandoc`)",
        "Or configure `pandoc_path` in `require('nvim-docx').setup({ pandoc_path = '/path/to/pandoc' })`",
      }
    )
  end

  -- Check Temp Directory
  local utils = require("nvim-docx.utils")
  local temp_dir = utils.ensure_temp_dir()
  if vim.fn.isdirectory(temp_dir) == 1 then
    vim.health.ok(string.format("Workspace cache directory writable: %s", temp_dir))
  else
    vim.health.warn(
      string.format("Failed to create workspace cache directory: %s", temp_dir),
      { "Check directory permissions for stdpath('cache')" }
    )
  end
end

return M
