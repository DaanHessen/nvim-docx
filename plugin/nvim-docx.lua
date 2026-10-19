if vim.g.loaded_nvim_docx == 1 then
  return
end
vim.g.loaded_nvim_docx = 1

-- Word documents of every generation: .docx/.docm/.dotx/.dotm (Word 2007+, OOXML)
-- and .doc/.dot (Word 97-2003 binary). Patterns are case-insensitive.
local docx_pattern = {
  "*.[dD][oO][cC][xX]",
  "*.[dD][oO][cC][mM]",
  "*.[dD][oO][tT][xX]",
  "*.[dD][oO][tT][mM]",
  "*.[dD][oO][cC]",
  "*.[dD][oO][tT]",
}

-- Register filetype with highest priority so bigfile plugins (e.g. snacks.bigfile) do not hijack Word files
if vim.filetype and vim.filetype.add then
  vim.filetype.add({
    pattern = {
      [".*%.[dD][oO][cCtT][xXmM]?$"] = { "markdown", { priority = math.huge } },
    },
  })
end

-- Neovim's bundled zipPlugin also claims *.docx/*.docm/*.dotx/*.dotm with its own BufReadCmd,
-- which runs alongside ours and marks the buffer readonly/nofile. Keep it away from Word files.
local word_zip_exts = { docx = true, docm = true, dotx = true, dotm = true }
if vim.g.zipPlugin_ext == nil then
  -- Same default list zipPlugin uses, minus the Word extensions
  vim.g.zipPlugin_ext = "*.aar,*.apk,*.celzip,*.crtx,*.ear,*.epub,*.gcsx,*.glox,*.gqsx,*.ja,*.jar,*.kmz,"
    .. "*.odb,*.odc,*.odf,*.odg,*.odi,*.odm,*.odp,*.ods,*.odt,*.otc,*.otf,*.otg,*.oth,*.oti,*.otp,*.ots,*.ott,"
    .. "*.oxt,*.pkpass,*.potm,*.potx,*.ppam,*.ppsm,*.ppsx,*.pptm,*.pptx,*.sldx,*.thmx,*.vdw,*.war,*.whl,*.wsz,"
    .. "*.xap,*.xlam,*.xlsb,*.xlsm,*.xlsx,*.xltm,*.xltx,*.xpi,*.zip"
end

local function detach_zip_plugin()
  local ok, autocmds = pcall(vim.api.nvim_get_autocmds, { group = "zip" })
  if not ok then
    return
  end
  for _, ac in ipairs(autocmds) do
    local ext = (ac.pattern or ""):match("^%*%.(%w+)$")
    if ext and word_zip_exts[ext:lower()] then
      -- :autocmd-defined entries have no id, so clear by group/event/pattern
      pcall(vim.api.nvim_clear_autocmds, { group = "zip", event = ac.event, pattern = ac.pattern })
    end
  end
end

detach_zip_plugin()

-- Create autocommand group
local augroup = vim.api.nvim_create_augroup("NvimDocx", { clear = true })

-- zipPlugin may be sourced after us (runtimepath order); detach it again once startup is done
vim.api.nvim_create_autocmd({ "VimEnter", "SourcePost" }, {
  group = augroup,
  callback = detach_zip_plugin,
})

-- Handle opening Word files using the custom reader
vim.api.nvim_create_autocmd("BufReadCmd", {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    require("nvim-docx").handle_docx_open(ev)
  end,
})

-- ":read file.docx" inserts the document as Markdown below the cursor
vim.api.nvim_create_autocmd("FileReadCmd", {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    require("nvim-docx").handle_docx_read(ev)
  end,
})

-- Fallback if a buffer was loaded raw before BufReadCmd was registered (e.g. lazy loading)
vim.api.nvim_create_autocmd("BufReadPost", {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    if not vim.b[ev.buf].nvim_docx_active then
      require("nvim-docx").handle_docx_open(ev)
    end
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

-- ":'<,'>write part.docx" compiles only the given lines into a new document
vim.api.nvim_create_autocmd("FileWriteCmd", {
  group = augroup,
  pattern = docx_pattern,
  callback = function(ev)
    require("nvim-docx").handle_docx_write_range(ev)
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
