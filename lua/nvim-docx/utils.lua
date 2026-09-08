local M = {}

local function get_options()
  local config = require("nvim-docx.config")
  return config.options or {}
end

local function join_path(lhs, rhs)
  if lhs == "" then
    return rhs
  end
  if lhs:sub(-1) == "/" then
    return lhs .. rhs
  end
  return lhs .. "/" .. rhs
end

local function unique_id()
  local hrtime = (vim.uv and vim.uv.hrtime) or (vim.loop and vim.loop.hrtime)
  if hrtime then
    return string.format("%x", hrtime())
  end
  return string.format("%x%x", os.time(), math.random(1000, 9999))
end

function M.is_docx_file(filepath)
  return vim.fn.fnamemodify(filepath, ":e"):lower() == "docx"
end

function M.get_temp_dir()
  local opts = get_options()
  return opts.temp_dir or (vim.fn.stdpath("cache") .. "/nvim-docx")
end

function M.ensure_dir(path)
  if vim.fn.isdirectory(path) == 0 then
    vim.fn.mkdir(path, "p")
  end
  return path
end

function M.ensure_temp_dir()
  return M.ensure_dir(M.get_temp_dir())
end

function M.create_workspace(docx_path)
  local base = M.ensure_temp_dir()
  local name = vim.fn.fnamemodify(docx_path or "docx", ":t:r")
  name = name:gsub("[^%w%-_%.]", "_")
  local workspace = join_path(base, string.format("%s-%s", name, unique_id()))
  return M.ensure_dir(workspace)
end

function M.cleanup_workspace(path)
  if path and path ~= "" and vim.fn.isdirectory(path) == 1 then
    vim.fn.delete(path, "rf")
  end
end

function M.media_dir_for(_, workspace)
  if workspace and workspace ~= "" then
    return workspace
  end
  return M.ensure_temp_dir()
end

function M.create_temp_md_file(root_dir)
  local base = root_dir or M.ensure_temp_dir()
  return join_path(base, string.format("docx-%s.md", unique_id()))
end

function M.cleanup_temp_file(filepath)
  if filepath and vim.fn.filereadable(filepath) == 1 then
    vim.fn.delete(filepath)
  end
end

function M.read_file_lines(path)
  local ok, result = pcall(vim.fn.readfile, path)
  if not ok then
    return nil, result
  end
  return result
end

function M.write_file_lines(path, lines)
  local ok, err = pcall(vim.fn.writefile, lines, path)
  if not ok then
    return false, err
  end
  return true
end

function M.notify(message, level)
  local lvl = vim.log.levels[string.upper(level or "info")] or vim.log.levels.INFO
  vim.notify(string.format("[nvim-docx] %s", message), lvl)
end

function M.debug(message)
  local opts = get_options()
  if opts.debug then
    vim.notify(string.format("[nvim-docx] %s", message), vim.log.levels.DEBUG)
  end
end

function M.escape_path(path)
  return vim.fn.shellescape(path)
end

return M
