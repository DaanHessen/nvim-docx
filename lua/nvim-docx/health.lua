local M = {}

-- vim.health.start/ok/... arrived in 0.9; 0.8 has the report_* names
local health = {
  start = vim.health.start or vim.health.report_start,
  ok = vim.health.ok or vim.health.report_ok,
  info = vim.health.info or vim.health.report_info,
  warn = vim.health.warn or vim.health.report_warn,
  error = vim.health.error or vim.health.report_error,
}

function M.check()
  health.start("nvim-docx")

  -- 1. Check Neovim version
  if vim.fn.has("nvim-0.8") == 1 then
    health.ok(string.format("Neovim version: %s", vim.version().string or ">= 0.8"))
  else
    health.error("Neovim 0.8+ is required")
  end

  -- 2. Check native pure-Lua OOXML engine
  local converter = require("nvim-docx.converter")
  local engine_ok, engine_info = converter.check_engine_available()
  if engine_ok then
    local jit_ver = ( _G.jit and _G.jit.version ) or _VERSION
    health.ok(string.format("Native Lua OOXML engine active (%s, zero external dependencies)", jit_ver))
  else
    health.error("Native Lua OOXML engine self-test failed: " .. tostring(engine_info))
  end

  -- Configs written for the pandoc-based versions keep working; the option is simply unused now
  local config = require("nvim-docx.config")
  if config.options.pandoc_path then
    health.info("`pandoc_path` is set but no longer used: nvim-docx no longer needs pandoc (safe to remove)")
  end

  -- 3. Check Temp Directory
  local utils = require("nvim-docx.utils")
  local temp_dir = utils.ensure_temp_dir()
  if vim.fn.isdirectory(temp_dir) == 1 then
    health.ok(string.format("Workspace cache directory writable: %s", temp_dir))
  else
    health.warn(
      string.format("Failed to create workspace cache directory: %s", temp_dir),
      { "Check directory permissions for stdpath('cache')" }
    )
  end
end

return M
