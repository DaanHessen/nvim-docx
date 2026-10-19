if vim.g.loaded_nvim_docx == 1 then
  return
end
vim.g.loaded_nvim_docx = 1

-- Create autocommand group
local augroup = vim.api.nvim_create_augroup("NvimDocx", { clear = true })

local docx_pattern = "*.[dD][oO][cC][xX]"

-- Handle opening .docx files using custom reader
vim.api.nvim_create_autocmd("BufReadCmd", {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    require("nvim-docx").handle_docx_open(ev)
  end,
})

-- Handle writing markdown back to DOCX
vim.api.nvim_create_autocmd("BufWriteCmd", {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    require("nvim-docx").handle_docx_save(ev)
  end,
})

-- Cleanup workspaces when buffers are deleted or wiped, or when Neovim exits
vim.api.nvim_create_autocmd({ "BufDelete", "BufWipeout" }, {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    require("nvim-docx").handle_docx_close(ev)
  end,
})

vim.api.nvim_create_autocmd("VimLeavePre", {
  group = augroup,
  callback = function()
    require("nvim-docx").cleanup_all_workspaces()
  end,
})

-- User commands for inspecting, exporting, and compiling DOCX files
vim.api.nvim_create_user_command("DocxInfo", function()
  require("nvim-docx").docx_info()
end, { desc = "Show document statistics and metadata for the current DOCX buffer" })

vim.api.nvim_create_user_command("DocxExportMarkdown", function(opts)
  require("nvim-docx").export_markdown(opts.args)
end, {
  nargs = "?",
  complete = "file",
  desc = "Export the current DOCX buffer and extracted media to a Markdown file",
})

vim.api.nvim_create_user_command("DocxFromMarkdown", function(opts)
  require("nvim-docx").from_markdown(opts.args)
end, {
  nargs = "?",
  complete = "file",
  desc = "Compile the current Markdown buffer directly to a .docx file",
})
