local M = {}

M.defaults = {
  temp_dir = vim.fn.stdpath("cache") .. "/nvim-docx", -- Base directory for workspaces
  markdown_format = "markdown",                        -- Markdown flavor for editing ("markdown", "gfm")
  auto_cleanup = true,                                 -- Wipe workspaces when the buffer is closed
  debug = false,                                       -- Enable debug logging
  preserve_styles = true,                              -- Preserve original .docx styles, headers, footers, and layout on save
  wrap = "none",                                       -- Text wrap mode ("none", "auto", "preserve")
  track_changes = nil,                                 -- Track changes handling ("accept", "reject", "all", or nil for default "accept")
  frontmatter = false,                                 -- Include YAML frontmatter for DOCX document properties
}

M.options = vim.deepcopy(M.defaults)

function M.setup(opts)
  -- Merge user options with defaults without mutating defaults table
  M.options = vim.tbl_deep_extend("force", vim.deepcopy(M.defaults), opts or {})

  -- Ensure temp directory exists
  local utils = require("nvim-docx.utils")
  utils.ensure_temp_dir()

  local converter = require("nvim-docx.converter")
  converter.invalidate_cache()

  if M.options.debug then
    utils.debug("Native Lua DOCX engine initialized (zero external dependencies)")
  end

  if vim.g.loaded_nvim_docx ~= 1 then
    pcall(vim.cmd, "runtime! plugin/nvim-docx.lua")
  end

  -- If any .docx buffer was already opened before setup() ran (e.g. via lazy-loading), initialize it now
  for _, bufnr in ipairs(vim.api.nvim_list_bufs()) do
    if vim.api.nvim_buf_is_loaded(bufnr) and not vim.b[bufnr].nvim_docx_active then
      local name = vim.api.nvim_buf_get_name(bufnr)
      if name ~= "" and utils.is_docx_file(name) then
        pcall(function()
          require("nvim-docx").handle_docx_open({ buf = bufnr, file = name })
        end)
      end
    end
  end

  return true
end

return M
