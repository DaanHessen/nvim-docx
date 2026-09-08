local utils = require("nvim-docx.utils")
local config = require("nvim-docx.config")

local M = {}

local pandoc_checked = false
local pandoc_available = false
local pandoc_version_str = ""

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

local function execute_command(args, cwd)
  utils.debug("Executing: " .. table.concat(args, " "))
  if vim.system then
    local ok_call, res = pcall(function()
      return vim.system(args, { cwd = cwd, text = true }):wait()
    end)
    if not ok_call or not res then
      return false, tostring(res or "failed to spawn process"), -1
    end
    local ok = (res.code == 0)
    local out = (res.stdout or "") .. (res.stderr or "")
    return ok, out, res.code
  else
    local escaped = {}
    for _, arg in ipairs(args) do
      table.insert(escaped, vim.fn.shellescape(arg))
    end
    local cmd_str = table.concat(escaped, " ")
    if cwd and cwd ~= "" then
      cmd_str = string.format("cd %s && %s 2>&1", vim.fn.shellescape(cwd), cmd_str)
    else
      cmd_str = cmd_str .. " 2>&1"
    end
    local output = vim.fn.system(cmd_str)
    local code = vim.v.shell_error
    return code == 0, output, code
  end
end

local function get_pandoc_bin()
  local opts = config.options or {}
  return opts.pandoc_path or "pandoc"
end

function M.invalidate_cache()
  pandoc_checked = false
  pandoc_available = false
  pandoc_version_str = ""
end

function M.check_pandoc_available(force)
  if not force and pandoc_checked then
    return pandoc_available, pandoc_version_str
  end

  pandoc_checked = true
  local bin = get_pandoc_bin()
  local ok, out = execute_command({ bin, "--version" })
  if ok and out ~= "" and not out:match("not found") then
    pandoc_available = true
    pandoc_version_str = out:match("pandoc%s+([%d%.]+)") or out:match("^([^\n]+)") or "available"
  else
    pandoc_available = false
    pandoc_version_str = ""
  end

  return pandoc_available, pandoc_version_str
end

local function ensure_pandoc()
  local available = M.check_pandoc_available()
  if available then
    return true
  end
  utils.notify("Pandoc is not available. Please install pandoc or configure `pandoc_path`.", "error")
  return false
end

function M.docx_to_markdown(docx_path, md_path, workspace)
  if not ensure_pandoc() then
    return false, "pandoc not available"
  end

  local opts = config.options or {}
  local format = opts.markdown_format or "markdown"
  local bin = get_pandoc_bin()

  -- Copy source docx into workspace as input.docx to prevent URI/path issues in pandoc across OS versions
  local safe_in = workspace .. "/input.docx"
  copy_file(docx_path, safe_in)

  local args = {
    bin,
    safe_in,
    "-f", "docx",
    "-t", format,
    "-o", md_path,
    "--extract-media=.",
  }

  if opts.wrap and opts.wrap ~= "" then
    table.insert(args, string.format("--wrap=%s", opts.wrap))
  end

  if opts.track_changes and opts.track_changes ~= "" then
    table.insert(args, string.format("--track-changes=%s", opts.track_changes))
  end

  local ok, output = execute_command(args, workspace)
  vim.fn.delete(safe_in)

  if not ok then
    local err_msg = "Pandoc conversion failed: " .. (output ~= "" and output or "unknown error")
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  return vim.fn.filereadable(md_path) == 1
end

function M.markdown_to_docx(md_path, docx_path, workspace, reference_doc)
  if not ensure_pandoc() then
    return false, "pandoc not available"
  end

  local opts = config.options or {}
  local format = opts.markdown_format or "markdown"
  local bin = get_pandoc_bin()
  local temp_out = workspace .. "/output_temp.docx"
  local safe_ref = workspace .. "/reference.docx"

  vim.fn.delete(temp_out)
  vim.fn.delete(safe_ref)

  local args = {
    bin,
    md_path,
    "-f", format,
    "-t", "docx",
    "--resource-path=.",
    "-o", temp_out,
  }

  local has_ref = false
  if opts.preserve_styles and reference_doc and reference_doc ~= "" and vim.fn.filereadable(reference_doc) == 1 then
    if copy_file(reference_doc, safe_ref) then
      table.insert(args, string.format("--reference-doc=%s", safe_ref))
      has_ref = true
    end
  end

  local ok, output = execute_command(args, workspace)
  if has_ref then
    vim.fn.delete(safe_ref)
  end

  if not ok then
    local err_msg = "Pandoc conversion to DOCX failed: " .. (output ~= "" and output or "unknown error")
    utils.notify(err_msg, "error")
    vim.fn.delete(temp_out)
    return false, err_msg
  end

  if vim.fn.filereadable(temp_out) ~= 1 then
    local err_msg = "Pandoc produced no output file: " .. temp_out
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  local replace_ok, replace_err = atomic_replace(temp_out, docx_path)
  if not replace_ok then
    local err_msg = "Failed to write target DOCX: " .. tostring(replace_err)
    utils.notify(err_msg, "error")
    return false, err_msg
  end

  return true
end

return M
