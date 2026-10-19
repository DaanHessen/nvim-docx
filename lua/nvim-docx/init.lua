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
  local name = vim.api.nvim_buf_get_name(bufnr)
  if name ~= "" and utils.is_docx_file(name) then
    return absolute_path(name)
  end
  if ev and ev.file and ev.file ~= "" and utils.is_docx_file(ev.file) then
    return absolute_path(ev.file)
  end
  if ev and ev.match and ev.match ~= "" and utils.is_docx_file(ev.match) then
    return absolute_path(ev.match)
  end
  return absolute_path(vim.fn.expand("%:p"))
end

local active_workspaces = {}

local function cleanup_state_files(state)
  if not state then
    return
  end
  if state.last_temp_file then
    utils.cleanup_temp_file(state.last_temp_file)
    state.last_temp_file = nil
  end
  if state.workspace then
    active_workspaces[state.workspace] = nil
    utils.cleanup_workspace(state.workspace)
    state.workspace = nil
  end
end

-- Full per-buffer state lives on the Lua side: it holds the loaded archive and parsed XML, which must
-- not be copied into a buffer variable. vim.b keeps a small summary for statuslines and scripts.
local buffer_states = {}

local function set_buffer_state(bufnr, state)
  buffer_states[bufnr] = state
  if state then
    if state.workspace then
      active_workspaces[state.workspace] = state
    end
    vim.b[bufnr].nvim_docx_state = {
      original_path = state.original_path,
      workspace = state.workspace,
      is_new = state.is_new,
      legacy = state.legacy,
    }
    vim.b[bufnr].nvim_docx_active = true
    vim.b[bufnr].nvim_docx_original_path = state.original_path
  else
    vim.b[bufnr].nvim_docx_state = nil
    vim.b[bufnr].nvim_docx_active = nil
    vim.b[bufnr].nvim_docx_original_path = nil
  end
end

local function get_buffer_state(bufnr)
  return buffer_states[bufnr]
end

local function write_buffer_lines(bufnr, lines)
  local old_ul = vim.bo[bufnr].undolevels
  vim.bo[bufnr].undolevels = -1
  vim.bo[bufnr].modifiable = true
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)
  vim.bo[bufnr].modified = false
  vim.bo[bufnr].undolevels = old_ul
end

function M.setup(opts)
  return config.setup(opts)
end

function M.handle_docx_open(ev)
  local bufnr = get_bufnr(ev)
  local prev_state = get_buffer_state(bufnr)
  local docx_path = resolve_docx_path(ev, bufnr)

  if (docx_path == "" or not utils.is_docx_file(docx_path)) and prev_state and prev_state.original_path then
    docx_path = prev_state.original_path
  end

  if docx_path == "" or not utils.is_docx_file(docx_path) then
    return
  end

  -- Clean up any previous workspace attached to this buffer (e.g., on :edit!)
  if prev_state and config.options.auto_cleanup then
    cleanup_state_files(prev_state)
  end

  local file_exists = (vim.fn.filereadable(docx_path) == 1)
  local workspace = utils.create_workspace(docx_path)
  local temp_md = utils.create_temp_md_file(workspace)
  local loaded_archive = nil
  local loaded_source = nil

  if file_exists then
    local ok, conv_err, archive, _, source = converter.docx_to_markdown(docx_path, temp_md, workspace)
    if not ok then
      if config.options.auto_cleanup then
        utils.cleanup_temp_file(temp_md)
        utils.cleanup_workspace(workspace)
      end
      error(conv_err or "Failed to open DOCX")
    end
    loaded_archive = archive
    loaded_source = source

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
  vim.bo[bufnr].readonly = false
  vim.bo[bufnr].swapfile = false
  vim.bo[bufnr].modified = false

  local state = {
    original_path = docx_path,
    workspace = workspace,
    is_new = not file_exists,
    -- Word 97-2003 binary: readable, but saves go to a .docx next to it
    legacy = file_exists and loaded_archive == nil,
    reference_archive = loaded_archive,
    preserve_source = loaded_source,
  }

  if config.options.auto_cleanup then
    utils.cleanup_temp_file(temp_md)
  else
    state.last_temp_file = temp_md
  end

  set_buffer_state(bufnr, state)
  vim.schedule(function()
    if vim.api.nvim_buf_is_valid(bufnr) and vim.b[bufnr].nvim_docx_active then
      if vim.bo[bufnr].filetype ~= "markdown" then
        vim.bo[bufnr].filetype = "markdown"
      end
    end
  end)
  if state.legacy then
    utils.notify(
      "Editing Word 97-2003 document as Markdown: " .. vim.fn.fnamemodify(docx_path, ":t")
        .. " (:w saves to " .. vim.fn.fnamemodify(docx_path, ":t:r") .. ".docx)",
      "info"
    )
  elseif file_exists then
    utils.notify("Editing DOCX as Markdown: " .. vim.fn.fnamemodify(docx_path, ":t"), "info")
  else
    utils.notify("New DOCX buffer (saves to: " .. vim.fn.fnamemodify(docx_path, ":t") .. ")", "info")
  end
end

local function determine_target_path(ev, bufnr, state)
  local buf_name = vim.api.nvim_buf_get_name(bufnr)
  local canonical_name = (buf_name ~= "") and absolute_path(buf_name) or ""
  local orig_path = (state and state.original_path) or ""

  -- Check if user issued an explicit "Save As" (:w different_name.docx)
  if ev and ev.file and ev.file ~= "" and ev.file ~= buf_name then
    local ev_abs = absolute_path(ev.file)
    if ev_abs ~= canonical_name and ev_abs ~= orig_path and utils.is_docx_file(ev_abs) and not ev.file:match("^[#%%]") then
      return ev_abs
    end
  end

  if orig_path ~= "" and utils.is_docx_file(orig_path) then
    return orig_path
  end

  if canonical_name ~= "" and utils.is_docx_file(canonical_name) then
    return canonical_name
  end

  if ev and ev.file and ev.file ~= "" then
    return absolute_path(ev.file)
  end

  return ""
end

function M.handle_docx_save(ev)
  local bufnr = get_bufnr(ev)
  local state = get_buffer_state(bufnr)

  if not state or not state.original_path then
    return
  end

  if not state.workspace or vim.fn.isdirectory(state.workspace) == 0 then
    state.workspace = utils.create_workspace(state.original_path)
    active_workspaces[state.workspace] = state
  end

  local target_path = determine_target_path(ev, bufnr, state)
  if target_path == "" then
    utils.notify("Unable to determine DOCX save path", "error")
    return
  end

  -- The binary Word 97-2003 format cannot be written; save as .docx next to it instead of
  -- overwriting the original with a file in a different format
  if utils.is_legacy_word_ext(target_path) then
    local docx_target = vim.fn.fnamemodify(target_path, ":r") .. ".docx"
    if vim.fn.filereadable(docx_target) == 1 and docx_target ~= state.legacy_target then
      utils.notify(
        "Not overwriting existing " .. vim.fn.fnamemodify(docx_target, ":t")
          .. "; save with :w <name>.docx",
        "error"
      )
      return
    end
    state.legacy_target = docx_target
    target_path = docx_target
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

  local save_ok, _, new_archive = converter.markdown_to_docx(temp_md, target_path, state.workspace, ref_doc, {
    reference_archive = (not state.is_new) and state.reference_archive or nil,
    preserve_source = (not state.is_new) and state.preserve_source or nil,
  })
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

  if target_path ~= state.original_path then
    utils.notify("Saved as " .. target_path, "info")
  end
  state.original_path = target_path
  state.is_new = false
  state.legacy = false
  if new_archive then
    state.reference_archive = new_archive
    -- Next save diffs against what was just written
    local ok_src, src = pcall(converter.source_from_archive, new_archive)
    state.preserve_source = ok_src and src or nil
  end
  set_buffer_state(bufnr, state)
  vim.api.nvim_buf_set_name(bufnr, target_path)
  vim.bo[bufnr].modified = false
  utils.notify("Saved DOCX: " .. vim.fn.fnamemodify(target_path, ":t"), "info")
end

-- ":read file.docx": insert the converted Markdown below the cursor line ('[ for :0read)
function M.handle_docx_read(ev)
  local bufnr = get_bufnr(ev)
  local path = absolute_path(ev and (ev.file ~= "" and ev.file or ev.match) or "")
  if vim.fn.filereadable(path) ~= 1 then
    utils.notify("Cannot read " .. path, "error")
    return
  end
  local state = get_buffer_state(bufnr)
  local ws = (state and state.workspace) or utils.create_workspace(path)
  local lines, err = converter.docx_to_markdown_lines(path, ws)
  if not state then
    utils.cleanup_workspace(ws)
  end
  if not lines then
    utils.notify("Failed to read " .. vim.fn.fnamemodify(path, ":t") .. ": " .. tostring(err), "error")
    return
  end
  local after = vim.fn.line("'[")
  if after < 0 then
    after = 0
  end
  vim.api.nvim_buf_set_lines(bufnr, after, after, false, lines)
end

-- ":'<,'>write part.docx": compile just those lines into a new document
function M.handle_docx_write_range(ev)
  local bufnr = get_bufnr(ev)
  local target = absolute_path(ev and (ev.file ~= "" and ev.file or ev.match) or "")
  if utils.is_legacy_word_ext(target) then
    target = vim.fn.fnamemodify(target, ":r") .. ".docx"
  end
  local first, last = vim.fn.line("'["), vim.fn.line("']")
  local lines = vim.api.nvim_buf_get_lines(bufnr, math.max(first - 1, 0), last, false)
  local state = get_buffer_state(bufnr)
  local cur_name = vim.api.nvim_buf_get_name(bufnr)
  local ws = utils.create_workspace(target)
  -- Images extracted from the source document live in its workspace
  local ok, err = converter.markdown_lines_to_docx(lines, target, ws, nil, {
    source_dir = (state and state.workspace) or (cur_name ~= "" and vim.fn.fnamemodify(cur_name, ":p:h")) or vim.fn.getcwd(),
  })
  utils.cleanup_workspace(ws)
  if ok then
    utils.notify(string.format("Wrote %d lines to %s", #lines, vim.fn.fnamemodify(target, ":t")), "info")
  else
    utils.notify("Failed to write " .. target .. ": " .. tostring(err), "error")
  end
end

function M.handle_docx_close(ev)
  local bufnr = get_bufnr(ev)
  local state = get_buffer_state(bufnr)
  if not state then
    return
  end

  if config.options.auto_cleanup then
    cleanup_state_files(state)
  end

  set_buffer_state(bufnr, nil)
end

function M.cleanup_all_workspaces()
  if not config.options.auto_cleanup then
    return
  end
  for ws, state in pairs(active_workspaces) do
    if state and state.last_temp_file then
      utils.cleanup_temp_file(state.last_temp_file)
    end
    utils.cleanup_workspace(ws)
    active_workspaces[ws] = nil
  end
end

function M.docx_info(bufnr)
  bufnr = bufnr or vim.api.nvim_get_current_buf()
  local state = get_buffer_state(bufnr)
  local docx_path = state and state.original_path or resolve_docx_path(nil, bufnr)

  local stats, err
  if state and (state.is_new or vim.bo[bufnr].modified or vim.fn.filereadable(docx_path) ~= 1) then
    local markdown = require("nvim-docx.markdown")
    local ooxml = require("nvim-docx.ooxml")
    local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
    local ast = markdown.markdown_to_ast(lines, config.options)
    stats = ooxml.get_docx_stats(state.reference_archive, ast)
  else
    if not docx_path or docx_path == "" or vim.fn.filereadable(docx_path) ~= 1 then
      utils.notify("Current buffer is not backed by an existing .docx file on disk", "warn")
      return nil
    end
    stats, err = converter.get_docx_info(docx_path)
    if not stats then
      utils.notify("Failed to read DOCX info: " .. tostring(err), "error")
      return nil
    end
  end

  local summary = {
    string.format("File: %s", vim.fn.fnamemodify(docx_path, ":t")),
    string.format("Words: %d | Characters: %d", stats.words, stats.chars),
    string.format(
      "Blocks: %d headings, %d paragraphs, %d list items, %d tables, %d images",
      stats.headings,
      stats.paragraphs,
      stats.list_items,
      stats.tables,
      stats.images
    ),
  }
  if stats.metadata and (stats.metadata.title or stats.metadata.author) then
    summary[#summary + 1] = string.format(
      "Metadata: title=%s, author=%s",
      tostring(stats.metadata.title or "-"),
      tostring(stats.metadata.author or "-")
    )
  end

  utils.notify(table.concat(summary, "\n"), "info")
  return stats
end

function M.export_markdown(target_md_path, bufnr)
  bufnr = bufnr or vim.api.nvim_get_current_buf()
  local state = get_buffer_state(bufnr)
  local base_docx = (state and state.original_path) or vim.api.nvim_buf_get_name(bufnr)

  if not target_md_path or target_md_path == "" then
    if base_docx == "" then
      utils.notify("Please specify a target Markdown path", "error")
      return false
    end
    target_md_path = vim.fn.fnamemodify(base_docx, ":p:r") .. ".md"
  else
    target_md_path = absolute_path(target_md_path)
  end

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local ok, err = utils.write_file_lines(target_md_path, lines)
  if not ok then
    utils.notify("Failed to export Markdown: " .. tostring(err), "error")
    return false
  end

  -- Media is only extracted on demand; make sure all of it is in the workspace before copying
  if state and state.workspace and state.reference_archive then
    utils.ensure_dir(state.workspace .. "/media")
    require("nvim-docx.ooxml").extract_media(state.reference_archive, state.workspace)
  end

  -- Copy extracted media directory next to target_md_path if present
  if state and state.workspace and vim.fn.isdirectory(state.workspace .. "/media") == 1 then
    local dst_media = vim.fn.fnamemodify(target_md_path, ":p:h") .. "/media"
    utils.ensure_dir(dst_media)
    local files = vim.fn.glob(state.workspace .. "/media/*", false, true)
    local uv = vim.uv or vim.loop
    for _, src_f in ipairs(files) do
      local dst_f = dst_media .. "/" .. vim.fn.fnamemodify(src_f, ":t")
      if uv and uv.fs_copyfile then
        uv.fs_copyfile(src_f, dst_f)
      end
    end
  end

  utils.notify("Exported Markdown to: " .. target_md_path, "info")
  return true, target_md_path
end

function M.from_markdown(target_docx_path, bufnr)
  bufnr = bufnr or vim.api.nvim_get_current_buf()
  local cur_name = vim.api.nvim_buf_get_name(bufnr)

  if not target_docx_path or target_docx_path == "" then
    if cur_name == "" then
      utils.notify("Please specify an output .docx path", "error")
      return false
    end
    target_docx_path = vim.fn.fnamemodify(cur_name, ":p:r") .. ".docx"
  else
    target_docx_path = absolute_path(target_docx_path)
    if not utils.is_docx_file(target_docx_path) then
      target_docx_path = target_docx_path .. ".docx"
    end
  end

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local source_dir = cur_name ~= "" and vim.fn.fnamemodify(cur_name, ":p:h") or vim.fn.getcwd()
  local ws = utils.create_workspace(target_docx_path)

  local ok, err = converter.markdown_lines_to_docx(lines, target_docx_path, ws, nil, {
    source_dir = source_dir,
  })
  utils.cleanup_workspace(ws)

  if not ok then
    return false, err
  end

  utils.notify("Compiled Markdown to DOCX: " .. target_docx_path, "info")
  return true, target_docx_path
end

return M
