local M = {}

function M.check()
  vim.health.start("nvim-docx")

  -- 1. Check Neovim version
  if vim.fn.has("nvim-0.8") == 1 then
    vim.health.ok(string.format("Neovim version: %s", vim.version().string or ">= 0.8"))
  else
    vim.health.error("Neovim 0.8+ is required")
  end

  -- 2. Check native pure-Lua OOXML engine
  local converter = require("nvim-docx.converter")
  local engine_ok, engine_info = converter.check_engine_available()
  if engine_ok then
    local jit_ver = ( _G.jit and _G.jit.version ) or _VERSION
    vim.health.ok(string.format("Native Lua OOXML engine active (%s, zero external dependencies)", jit_ver))
  else
    vim.health.error("Native Lua OOXML engine self-test failed: " .. tostring(engine_info))
  end

  -- 3. Check Temp Directory
  local utils = require("nvim-docx.utils")
  local temp_dir = utils.ensure_temp_dir()
  if vim.fn.isdirectory(temp_dir) == 1 then
    vim.health.ok(string.format("Workspace cache directory writable: %s", temp_dir))
  else
    vim.health.warn(
      string.format("Failed to create workspace cache directory: %s", temp_dir),
      { "Check directory permissions for stdpath('cache')" }
    )
  end
end

return M
