local converter = require("nvim-docx.converter")
local config = require("nvim-docx.config")
local utils = require("nvim-docx.utils")

local M = {}

local function get_bufnr(ev)
  if ev and ev.buf then
    return ev.buf
  end
  return vim.api.nvim_get_current_buf()
end

local function absolute_path(path)
  if not path or path == "" then
    return ""
  end
  return vim.fn.fnamemodify(path, ":p")
end

local function resolve_docx_path(ev, bufnr)
  if ev and ev.file and ev.file ~= "" then
    return absolute_path(ev.file)
  end
  if ev and ev.match and ev.match ~= "" then
    return absolute_path(ev.match)
  end
  local name = vim.api.nvim_buf_get_name(bufnr)
  if name ~= "" then
    return absolute_path(name)
  end
  return absolute_path(vim.fn.expand("%:p"))
end

local function set_buffer_state(bufnr, state)
  if state then
    vim.b[bufnr].nvim_docx_state = state
    vim.b[bufnr].nvim_docx_active = true
    vim.b[bufnr].nvim_docx_original_path = state.original_path
  else
    vim.b[bufnr].nvim_docx_state = nil
    vim.b[bufnr].nvim_docx_active = nil
    vim.b[bufnr].nvim_docx_original_path = nil
  end
end

local function get_buffer_state(bufnr)
  return vim.b[bufnr].nvim_docx_state
end

local function write_buffer_lines(bufnr, lines)
  vim.bo[bufnr].modifiable = true
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)
  vim.bo[bufnr].modified = false
end

function M.setup(opts)
  return config.setup(opts)
end

function M.handle_docx_open(ev)
  local bufnr = get_bufnr(ev)
  local docx_path = resolve_docx_path(ev, bufnr)

  if docx_path == "" or not utils.is_docx_file(docx_path) then
    return
  end

  local file_exists = (vim.fn.filereadable(docx_path) == 1)
  local workspace = utils.create_workspace(docx_path)
  local temp_md = utils.create_temp_md_file(workspace)

  if file_exists then
    local ok, conv_err = converter.docx_to_markdown(docx_path, temp_md, workspace)
    if not ok then
      if config.options.auto_cleanup then
        utils.cleanup_temp_file(temp_md)
        utils.cleanup_workspace(workspace)
      end
      return
    end

    local content, err = utils.read_file_lines(temp_md)
    if not content then
      utils.notify("Failed to read converted markdown: " .. tostring(err), "error")
      if config.options.auto_cleanup then
        utils.cleanup_temp_file(temp_md)
        utils.cleanup_workspace(workspace)
      end
      return
    end

    write_buffer_lines(bufnr, content)
  else
    -- Creating a new DOCX file
    write_buffer_lines(bufnr, { "" })
  end

  vim.bo[bufnr].filetype = "markdown"
  vim.bo[bufnr].buftype = "acwrite"
  vim.bo[bufnr].modified = false

  local state = {
    original_path = docx_path,
    workspace = workspace,
    is_new = not file_exists,
  }

  if config.options.auto_cleanup then
    utils.cleanup_temp_file(temp_md)
  else
    state.last_temp_file = temp_md
  end

  set_buffer_state(bufnr, state)
  if file_exists then
    utils.notify("Editing DOCX as Markdown: " .. vim.fn.fnamemodify(docx_path, ":t"), "info")
  else
    utils.notify("New DOCX buffer (saves to: " .. vim.fn.fnamemodify(docx_path, ":t") .. ")", "info")
  end
end

local function determine_target_path(ev, bufnr, state)
  if ev and ev.file and ev.file ~= "" then
    return absolute_path(ev.file)
  end
  local name = vim.api.nvim_buf_get_name(bufnr)
  if name ~= "" then
    return absolute_path(name)
  end
  return state.original_path
end

function M.handle_docx_save(ev)
  local bufnr = get_bufnr(ev)
  local state = get_buffer_state(bufnr)

  if not state or not state.original_path then
    return
  end

  local target_path = determine_target_path(ev, bufnr, state)
  if target_path == "" then
    utils.notify("Unable to determine DOCX save path", "error")
    return
  end

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local temp_md = utils.create_temp_md_file(state.workspace)
  local ok, write_err = utils.write_file_lines(temp_md, lines)
  if not ok then
    utils.notify("Failed to write temporary markdown: " .. tostring(write_err), "error")
    return
  end

  local ref_doc = state.original_path
  if state.is_new or vim.fn.filereadable(ref_doc) ~= 1 then
    ref_doc = nil
  end

  local save_ok, save_err = converter.markdown_to_docx(temp_md, target_path, state.workspace, ref_doc)
  if not save_ok then
    if config.options.auto_cleanup then
      utils.cleanup_temp_file(temp_md)
    end
    return
  end

  if config.options.auto_cleanup then
    utils.cleanup_temp_file(temp_md)
  else
    state.last_temp_file = temp_md
  end

  state.original_path = target_path
  state.is_new = false
  set_buffer_state(bufnr, state)
  vim.api.nvim_buf_set_name(bufnr, target_path)
  vim.bo[bufnr].modified = false
  utils.notify("Saved DOCX: " .. vim.fn.fnamemodify(target_path, ":t"), "info")
end

function M.handle_docx_close(ev)
  local bufnr = get_bufnr(ev)
  local state = get_buffer_state(bufnr)
  if not state then
    return
  end

  if config.options.auto_cleanup then
    if state.last_temp_file then
      utils.cleanup_temp_file(state.last_temp_file)
    end
    if state.workspace then
      utils.cleanup_workspace(state.workspace)
    end
  end

  set_buffer_state(bufnr, nil)
end

return M
