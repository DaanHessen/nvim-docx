if vim.g.loaded_nvim_docx == 1 then
  return
end
vim.g.loaded_nvim_docx = 1

-- Create autocommand group
local augroup = vim.api.nvim_create_augroup("NvimDocx", { clear = true })

-- Handle opening .docx files using custom reader
vim.api.nvim_create_autocmd("BufReadCmd", {
  group = augroup,
  pattern = "*.docx",
  callback = function(ev)
    require("nvim-docx").handle_docx_open(ev)
  end,
})

-- Handle writing markdown back to DOCX
vim.api.nvim_create_autocmd("BufWriteCmd", {
  group = augroup,
  pattern = "*.docx",
  callback = function(ev)
    require("nvim-docx").handle_docx_save(ev)
  end,
})

-- Cleanup workspaces when buffers are wiped
vim.api.nvim_create_autocmd("BufWipeout", {
  group = augroup,
  pattern = "*.docx",
  callback = function(ev)
    require("nvim-docx").handle_docx_close(ev)
  end,
})
