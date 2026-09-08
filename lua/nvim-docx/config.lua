local M = {}

M.defaults = {
  pandoc_path = "pandoc",                              -- Path to pandoc executable
  temp_dir = vim.fn.stdpath("cache") .. "/nvim-docx", -- Base directory for workspaces
  markdown_format = "markdown",                        -- Markdown flavor for editing ("markdown", "gfm", etc.)
  auto_cleanup = true,                                 -- Wipe workspaces when the buffer is closed
  debug = false,                                       -- Enable debug logging
  preserve_styles = true,                              -- Use original .docx as reference-doc on save to preserve styling
  wrap = "none",                                       -- Text wrap mode for pandoc ("none", "auto", "preserve")
  track_changes = nil,                                 -- Track changes handling ("accept", "reject", "all", or nil)
}

M.options = vim.deepcopy(M.defaults)

function M.setup(opts)
  -- Merge user options with defaults without mutating defaults table
  M.options = vim.tbl_deep_extend("force", vim.deepcopy(M.defaults), opts or {})

  -- Ensure temp directory exists
  local utils = require("nvim-docx.utils")
  utils.ensure_temp_dir()

  -- Invalidate pandoc cache in case pandoc_path changed
  local converter = require("nvim-docx.converter")
  converter.invalidate_cache()

  if M.options.debug then
    local available, ver = converter.check_pandoc_available()
    if not available then
      utils.notify("Pandoc not found at: " .. tostring(M.options.pandoc_path), "warn")
    else
      utils.debug("Pandoc detected: " .. tostring(ver))
    end
  end

  return true
end

return M
