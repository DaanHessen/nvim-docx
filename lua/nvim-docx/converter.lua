local utils = require("nvim-docx.utils")
local config = require("nvim-docx.config")
local zip = require("nvim-docx.zip")
local ooxml = require("nvim-docx.ooxml")
local doc_ole = require("nvim-docx.doc_ole")
local markdown = require("nvim-docx.markdown")
local preserve = require("nvim-docx.preserve")

local M = {}

local function copy_file(src, dst)
  local uv = vim.uv or vim.loop
  if uv and uv.fs_copyfile then
    local ok, err = uv.fs_copyfile(src, dst)
    return ok, err
  end
  local in_file = io.open(src, "rb")
  if not in_file then return false, "cannot open src" end
  local content = in_file:read("*a")
  in_file:close()
  local out_file = io.open(dst, "wb")
  if not out_file then return false, "cannot open dst" end
  out_file:write(content)
  out_file:close()
  return true
end

local function atomic_replace(src, dst)
  if vim.fn.rename(src, dst) == 0 then
    return true
  end
  local ok, err = copy_file(src, dst)
  if ok then
    vim.fn.delete(src)
    return true
  end
  return false, err
end

function M.invalidate_cache()
  -- Retained for API compatibility
end

function M.check_engine_available()
  local sample = "nvim-docx-self-test"
  local ok, inflated = pcall(function()
    return zip.inflate(zip.deflate(sample))
  end)
  if ok and inflated == sample then
    return true, "native-lua"
  end
  return false, tostring(inflated)
end

-- Kept for configs and scripts written against the pandoc-based versions: there is no external
-- dependency any more, so this reports the built-in engine instead
function M.check_pandoc_available()
  return M.check_engine_available()
end

function M.docx_to_markdown_lines(docx_path, workspace)
  local opts = config.options or {}
  utils.debug("Reading DOCX/DOC archive: " .. tostring(docx_path))

  if doc_ole.is_legacy_word_file(docx_path) then
    local ast, ast_err = doc_ole.doc_to_ast(docx_path, workspace)
    if not ast then
      local err_msg = "Word document conversion failed: " .. tostring(ast_err)
      utils.notify(err_msg, "error")
      return nil, err_msg
    end
    local lines = markdown.ast_to_markdown_lines(ast, opts)
    return lines, nil, nil, ast
  end

  local archive, zip_err = zip.read_zip(docx_path)
  if not archive then
    local err_msg = "DOCX conversion failed: " .. tostring(zip_err)
    utils.notify(err_msg, "error")
    return nil, err_msg
  end

  local ast, ast_err = ooxml.docx_to_ast(archive, workspace, opts)
  if not ast then
    local err_msg = "DOCX conversion failed: " .. tostring(ast_err)
    utils.notify(err_msg, "error")
    return nil, err_msg
  end

  local ranges = {}
  local lines = markdown.ast_to_markdown_lines(ast, vim.tbl_extend("force", opts, { block_ranges = ranges }))
  return lines, nil, archive, ast, preserve.make_source(lines, ranges, ast)
end

-- Rebuild the block mapping source from an archive (after a save, without extracting media again)
function M.source_from_archive(archive)
  local opts = config.options or {}
  local ast = ooxml.docx_to_ast(archive, nil, opts)
  if not ast then
    return nil
  end
  local ranges = {}
  local lines = markdown.ast_to_markdown_lines(ast, vim.tbl_extend("force", opts, { block_ranges = ranges }))
  return preserve.make_source(lines, ranges, ast)
end

function M.docx_to_markdown(docx_path, md_path, workspace)
  local lines, err, archive, ast, source = M.docx_to_markdown_lines(docx_path, workspace)
  if not lines then
    return false, err
  end

  local write_ok, write_err = utils.write_file_lines(md_path, lines)
  if not write_ok then
    local err_msg = "Failed to write converted Markdown: " .. tostring(write_err)
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  return vim.fn.filereadable(md_path) == 1, nil, archive, ast, source
end

function M.markdown_lines_to_docx(lines, docx_path, workspace, reference_doc, extra_opts)
  extra_opts = extra_opts or {}
  -- Large objects are passed by reference, never deep-copied into the merged options
  local ref_archive = extra_opts.reference_archive
  local source = extra_opts.preserve_source
  local plain_extra = vim.tbl_extend("force", {}, extra_opts)
  plain_extra.reference_archive = nil
  plain_extra.preserve_source = nil
  local opts = vim.tbl_deep_extend("force", config.options or {}, plain_extra)
  local ws = workspace or utils.ensure_temp_dir()
  local temp_out = ws .. "/output_temp.docx"

  vim.fn.delete(temp_out)

  local new_ranges = {}
  local ok_parse, ast = pcall(markdown.markdown_to_ast, lines, vim.tbl_extend("force", opts, { block_ranges = new_ranges }))
  if not ok_parse or not ast then
    local err_msg = "Markdown parsing failed: " .. tostring(ast)
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  if not ref_archive and opts.preserve_styles and reference_doc and reference_doc ~= "" and vim.fn.filereadable(reference_doc) == 1 and not doc_ole.is_legacy_word_file(reference_doc) then
    ref_archive = zip.read_zip(reference_doc)
  elseif not opts.preserve_styles then
    ref_archive = nil
  end

  if ref_archive and source then
    local ok_map, map = pcall(preserve.map_blocks, source, lines, new_ranges, ast.blocks)
    if ok_map then
      opts.preserve_source = source
      opts.preserve_map = map
    else
      utils.debug("Block mapping failed, regenerating whole document: " .. tostring(map))
    end
  end

  local ok_compile, new_archive = pcall(ooxml.ast_to_docx, ast, ws, ref_archive, opts)
  if not ok_compile or not new_archive then
    local err_msg = "DOCX compilation failed: " .. tostring(new_archive)
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  local write_ok, write_err = zip.write_zip(temp_out, new_archive)
  if not write_ok then
    local err_msg = "Failed to write temporary DOCX: " .. tostring(write_err)
    utils.notify(err_msg, "error")
    vim.fn.delete(temp_out)
    return false, err_msg
  end

  if vim.fn.filereadable(temp_out) ~= 1 then
    local err_msg = "DOCX engine produced no output file: " .. temp_out
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  local dst_dir = vim.fn.fnamemodify(docx_path, ":p:h")
  if dst_dir and dst_dir ~= "" and vim.fn.isdirectory(dst_dir) == 0 then
    vim.fn.mkdir(dst_dir, "p")
  end

  local replace_ok, replace_err = atomic_replace(temp_out, docx_path)
  if not replace_ok then
    local err_msg = "Failed to write target DOCX: " .. tostring(replace_err)
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  return true, nil, new_archive
end

function M.markdown_to_docx(md_path, docx_path, workspace, reference_doc, extra_opts)
  local lines, read_err = utils.read_file_lines(md_path)
  if not lines then
    local err_msg = "Failed to read Markdown file: " .. tostring(read_err)
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  local source_dir = vim.fn.fnamemodify(md_path, ":p:h")
  local merged_opts = vim.tbl_extend("force", { source_dir = source_dir }, extra_opts or {})
  return M.markdown_lines_to_docx(lines, docx_path, workspace, reference_doc, merged_opts)
end

function M.get_docx_info(docx_path)
  if doc_ole.is_legacy_word_file(docx_path) then
    local ast, ast_err = doc_ole.doc_to_ast(docx_path, nil)
    if not ast then
      return nil, ast_err
    end
    return ooxml.get_docx_stats(nil, ast)
  end
  local archive, err = zip.read_zip(docx_path)
  if not archive then
    return nil, err
  end
  local ast, ast_err = ooxml.docx_to_ast(archive, nil, config.options)
  if not ast then
    return nil, ast_err
  end
  return ooxml.get_docx_stats(archive, ast)
end

return M
