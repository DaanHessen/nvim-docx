vim.cmd("runtime! plugin/nvim-docx.lua")

local function assert_true(cond, msg)
  if not cond then
    error("Assertion failed: " .. (msg or "condition was false"))
  end
end

local function assert_equal(actual, expected, msg)
  if actual ~= expected then
    error(string.format("Assertion failed: %s (expected: %s, got: %s)", msg or "", tostring(expected), tostring(actual)))
  end
end

local tests = {}
local test_dir = "/tmp/nvim-docx-e2e-tests"

local function setup_env()
  vim.fn.delete(test_dir, "rf")
  vim.fn.mkdir(test_dir, "p")
  vim.fn.mkdir(test_dir .. "/assets", "p")

  local python_cmd = string.format([[
python3 -c "
from PIL import Image
img = Image.new('RGB', (80, 80), color=(100, 150, 200))
img.save('%s/assets/logo.png')
"
]], test_dir)
  vim.fn.system(python_cmd)

  local md_content = [[
# Test Document Header

This is a paragraph with **bold**, *italic*, and a [hyperlink](https://github.com/DaanHessen/nvim-docx).

## Section Two: Lists and Tables

Here is a list:
- Item Alpha
- Item Beta
  - Subitem Beta.1
- Item Gamma

| Col 1 | Col 2 | Col 3 |
| :--- | :---: | ---: |
| Left | Center | Right |
| Data A | Data B | Data C |

Here is an image:

![Test Logo](assets/logo.png)

Conclusion paragraph.
]]
  local md_file = test_dir .. "/source.md"
  vim.fn.writefile(vim.split(md_content, "\n"), md_file)

  local gen_cmd = string.format(
    "cd %s && pandoc source.md -o %s/rich.docx",
    vim.fn.shellescape(test_dir),
    vim.fn.shellescape(test_dir)
  )
  vim.fn.system(gen_cmd)
end

-- Test 1: Config and backward compatibility
tests["config_defaults_and_setup"] = function()
  local config = require("nvim-docx.config")
  config.setup({
    pandoc_path = "pandoc",
    debug = false,
    auto_cleanup = true,
  })
  assert_equal(config.options.pandoc_path, "pandoc", "pandoc_path")
  assert_equal(config.options.debug, false, "debug")
  assert_equal(config.options.auto_cleanup, true, "auto_cleanup")
  assert_equal(config.options.preserve_styles, true, "preserve_styles default")
  assert_equal(config.options.wrap, "none", "wrap default")

  config.setup()
  assert_true(config.options ~= nil, "options exist after empty setup")
end

-- Test 2: Checkhealth
tests["checkhealth_runs_without_error"] = function()
  local health = require("nvim-docx.health")
  local ok, err = pcall(health.check)
  assert_true(ok, "health.check succeeded: " .. tostring(err))
end

-- Test 3: Open rich docx, verify markdown buffer, edit, save, verify docx
tests["open_edit_save_rich_docx"] = function()
  local rich_docx = test_dir .. "/rich.docx"
  assert_true(vim.fn.filereadable(rich_docx) == 1, "rich.docx exists")

  vim.cmd("edit " .. vim.fn.fnameescape(rich_docx))
  local bufnr = vim.api.nvim_get_current_buf()

  assert_equal(vim.bo[bufnr].filetype, "markdown", "filetype is markdown")
  assert_equal(vim.bo[bufnr].buftype, "acwrite", "buftype is acwrite")
  assert_equal(vim.bo[bufnr].modified, false, "buffer not modified upon open")

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local text = table.concat(lines, "\n")
  assert_true(text:find("Test Document Header") ~= nil, "Contains heading")
  assert_true(text:find("Item Alpha") ~= nil, "Contains list item")
  assert_true(text:find("Col 1") ~= nil, "Contains table column")
  assert_true(text:find("logo%.png") ~= nil or text:find("media") ~= nil, "Contains image reference")

  table.insert(lines, "Added by automated test verification.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)
  assert_true(vim.bo[bufnr].modified, "buffer is marked modified")

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "buffer marked unmodified after write")

  local unzip_check = vim.fn.system("unzip -l " .. vim.fn.shellescape(rich_docx))
  assert_true(unzip_check:find("media") ~= nil, "DOCX still contains media archive: " .. unzip_check)

  local check_md = test_dir .. "/check.md"
  vim.fn.system(string.format("pandoc %s --wrap=none -o %s", vim.fn.shellescape(rich_docx), vim.fn.shellescape(check_md)))
  local check_text = table.concat(vim.fn.readfile(check_md), "\n")
  assert_true(check_text:find("Added by automated test verification") ~= nil, "Saved docx contains edited text")

  vim.cmd("bdelete!")
end

-- Test 4: Real Microsoft Word 2007+ document
tests["real_word_document_roundtrip"] = function()
  if vim.fn.filereadable("/tmp/sample_real.docx") ~= 1 then
    print("Skipping real word doc test (file not found)")
    return
  end

  local sample_copy = test_dir .. "/real_sample_test.docx"
  vim.fn.system(string.format("cp /tmp/sample_real.docx %s", vim.fn.shellescape(sample_copy)))

  vim.cmd("edit " .. vim.fn.fnameescape(sample_copy))
  local bufnr = vim.api.nvim_get_current_buf()

  assert_equal(vim.bo[bufnr].filetype, "markdown", "filetype is markdown")
  assert_equal(vim.bo[bufnr].buftype, "acwrite", "buftype is acwrite")

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  table.insert(lines, "# Additional Heading by Test")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "modified is false after write")

  local check_out = vim.fn.system("pandoc " .. vim.fn.shellescape(sample_copy) .. " -t plain")
  assert_true(check_out:find("Additional Heading by Test") ~= nil, "Pandoc can read modified real docx")

  vim.cmd("bdelete!")
end

-- Test 5: Special characters in filename
tests["filename_with_spaces_and_symbols"] = function()
  local tricky_name = test_dir .. "/Report - 2026 & Test (Special) #1.docx"
  vim.fn.system(string.format("cp %s/rich.docx %s", vim.fn.shellescape(test_dir), vim.fn.shellescape(tricky_name)))

  vim.cmd("edit " .. vim.fn.fnameescape(tricky_name))
  local bufnr = vim.api.nvim_get_current_buf()

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  table.insert(lines, "Tricky filename test line.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "saved tricky filename")

  local check_md = test_dir .. "/tricky_check.md"
  vim.fn.system(string.format("pandoc %s --wrap=none -o %s", vim.fn.shellescape(tricky_name), vim.fn.shellescape(check_md)))
  local check_text = table.concat(vim.fn.readfile(check_md), "\n")
  assert_true(check_text:find("Tricky filename test line") ~= nil, "Content persisted in tricky filename")

  vim.cmd("bdelete!")
end

-- Test 6: Creating a brand new docx from scratch
tests["create_new_docx_from_scratch"] = function()
  local new_docx = test_dir .. "/created_from_scratch.docx"
  vim.fn.delete(new_docx)

  vim.cmd("edit " .. vim.fn.fnameescape(new_docx))
  local bufnr = vim.api.nvim_get_current_buf()

  assert_equal(vim.bo[bufnr].filetype, "markdown", "filetype is markdown")
  assert_equal(vim.bo[bufnr].buftype, "acwrite", "buftype is acwrite")

  local new_lines = {
    "# Brand New Document",
    "",
    "Created directly from within Neovim!",
    "",
    "- Item 1",
    "- Item 2",
  }
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, new_lines)
  vim.cmd("write")

  assert_true(vim.fn.filereadable(new_docx) == 1, "New DOCX was created on disk")

  local check_out = vim.fn.system("pandoc " .. vim.fn.shellescape(new_docx) .. " -t plain")
  assert_true(check_out:find("Brand New Document") ~= nil, "New DOCX contains markdown content")

  vim.cmd("bdelete!")
end

-- Test 7: Save as new file (:w other_name.docx)
tests["save_as_new_filename"] = function()
  local orig_docx = test_dir .. "/rich.docx"
  local copy_docx = test_dir .. "/copy_via_saveas.docx"
  vim.fn.delete(copy_docx)

  vim.cmd("edit " .. vim.fn.fnameescape(orig_docx))
  local bufnr = vim.api.nvim_get_current_buf()

  vim.cmd("write " .. vim.fn.fnameescape(copy_docx))
  assert_true(vim.fn.filereadable(copy_docx) == 1, "copy_via_saveas.docx was created")

  local check_out = vim.fn.system("pandoc " .. vim.fn.shellescape(copy_docx) .. " -t plain")
  assert_true(check_out:find("Test Document Header") ~= nil, "Copy contains original content")

  vim.cmd("bdelete!")
end

-- Test 8: Workspace cleanup on buffer wipeout
tests["workspace_lifecycle_cleanup"] = function()
  local config = require("nvim-docx.config")
  config.setup({ auto_cleanup = true })

  local docx = test_dir .. "/rich.docx"
  vim.cmd("edit " .. vim.fn.fnameescape(docx))
  local bufnr = vim.api.nvim_get_current_buf()

  local state = vim.b[bufnr].nvim_docx_state
  assert_true(state ~= nil, "state exists")
  local workspace = state.workspace
  assert_true(vim.fn.isdirectory(workspace) == 1, "workspace directory exists while editing")

  vim.cmd("bwipeout " .. bufnr)
  assert_equal(vim.fn.isdirectory(workspace), 0, "workspace deleted after bwipeout with auto_cleanup=true")
end

-- Test 9: Corrupted docx file handling
tests["corrupted_docx_handling"] = function()
  local corrupt_docx = test_dir .. "/corrupted.docx"
  vim.fn.writefile({ "This is not a zip or docx file at all!" }, corrupt_docx)

  local ok, err = pcall(function()
    vim.cmd("edit " .. vim.fn.fnameescape(corrupt_docx))
  end)
  assert_true(not ok, "Opening corrupted file correctly reports error")
  assert_true(tostring(err):find("Pandoc") ~= nil, "Error mentions Pandoc: " .. tostring(err))
  vim.cmd("bdelete!")
end

-- Test 10: Invalid pandoc binary path handling
tests["invalid_pandoc_path_handling"] = function()
  local config = require("nvim-docx.config")
  config.setup({ pandoc_path = "/nonexistent/path/to/pandoc" })

  local docx = test_dir .. "/rich.docx"
  local ok, err = pcall(function()
    vim.cmd("edit " .. vim.fn.fnameescape(docx))
  end)

  -- Always restore valid configuration
  config.setup({ pandoc_path = "pandoc" })
  pcall(vim.cmd, "bdelete!")

  assert_true(not ok, "Invalid pandoc path correctly reports error")
  assert_true(tostring(err):find("[Pp]andoc") ~= nil, "Error mentions Pandoc: " .. tostring(err))
end

-- Main runner
setup_env()

local failed = 0
local passed = 0

print("\n================== RUNNING NVIM-DOCX TEST SUITE ==================")
for name, fn in pairs(tests) do
  local ok, err = pcall(fn)
  if ok then
    passed = passed + 1
    print("  ✓ PASS: " .. name)
  else
    failed = failed + 1
    print("  ✗ FAIL: " .. name .. " -> " .. tostring(err))
  end
end

print("==================================================================")
print(string.format("Result: %d passed, %d failed\n", passed, failed))

if failed > 0 then
  vim.cmd("cq")
else
  vim.cmd("qall!")
end
