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
local fixtures_dir = vim.fn.fnamemodify(debug.getinfo(1).source:sub(2), ":p:h") .. "/fixtures"

local function setup_env()
  vim.fn.delete(test_dir, "rf")
  vim.fn.mkdir(test_dir, "p")
  vim.fn.mkdir(test_dir .. "/assets", "p")

  -- Generate valid PNG using pure Python standard library (no pip / Pillow dependency)
  local python_cmd = string.format([[
python3 -c "
import struct, zlib
sig = b'\x89PNG\r\n\x1a\n'
ihdr_data = struct.pack('>IIBBBBB', 16, 16, 8, 2, 0, 0, 0)
ihdr = struct.pack('>I', len(ihdr_data)) + b'IHDR' + ihdr_data + struct.pack('>I', zlib.crc32(b'IHDR' + ihdr_data) & 0xffffffff)
raw = b''.join(b'\x00' + b'\x1f\x4a\x8c' * 16 for _ in range(16))
comp = zlib.compress(raw)
idat = struct.pack('>I', len(comp)) + b'IDAT' + comp + struct.pack('>I', zlib.crc32(b'IDAT' + comp) & 0xffffffff)
iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', zlib.crc32(b'IEND') & 0xffffffff)
with open('%s/assets/logo.png', 'wb') as f:
    f.write(sig + ihdr + idat + iend)
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

-- Test 4: Physical Microsoft Word document with embedded images
tests["physical_docx_with_embedded_images"] = function()
  local src_fixture = fixtures_dir .. "/having-images.docx"
  assert_true(vim.fn.filereadable(src_fixture) == 1, "having-images.docx fixture exists")

  local test_copy = test_dir .. "/test_having_images.docx"
  vim.fn.system(string.format("cp %s %s", vim.fn.shellescape(src_fixture), vim.fn.shellescape(test_copy)))

  -- Initial check: verify source fixture has media
  local initial_unzip = vim.fn.system("unzip -l " .. vim.fn.shellescape(test_copy))
  assert_true(initial_unzip:find("image1%.png") ~= nil, "Fixture contains image1.png")
  assert_true(initial_unzip:find("image2%.png") ~= nil, "Fixture contains image2.png")

  -- Open in Neovim
  vim.cmd("edit " .. vim.fn.fnameescape(test_copy))
  local bufnr = vim.api.nvim_get_current_buf()

  assert_equal(vim.bo[bufnr].filetype, "markdown", "filetype is markdown")
  assert_equal(vim.bo[bufnr].buftype, "acwrite", "buftype is acwrite")

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)

  -- Verify extracted media images are referenced in Markdown
  local text = table.concat(lines, "\n")
  assert_true(text:find("media") ~= nil, "Extracted media referenced in markdown buffer: " .. text)

  -- Edit buffer: add custom heading and text
  table.insert(lines, 1, "# Customized Heading in Image Document")
  table.insert(lines, 2, "Verified that images remain intact after editing.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  -- Save
  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "modified is false after saving")

  -- Verify images are STILL present in the saved DOCX zip archive
  local saved_unzip = vim.fn.system("unzip -l " .. vim.fn.shellescape(test_copy))
  assert_true(saved_unzip:find("word/media/image") ~= nil, "Media folder preserved in saved docx: " .. saved_unzip)

  -- Re-open in Neovim to verify roundtrip fidelity
  vim.cmd("edit!")
  local reloaded_lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local reloaded_text = table.concat(reloaded_lines, "\n")
  assert_true(reloaded_text:find("Customized Heading in Image Document") ~= nil, "Reloaded buffer has edited heading")
  assert_true(reloaded_text:find("media") ~= nil, "Reloaded buffer retains image references")

  vim.cmd("bdelete!")
end

-- Test 5: Physical Microsoft Word document with custom styles and fonts
tests["physical_docx_styled_document"] = function()
  local src_fixture = fixtures_dir .. "/styled-test.docx"
  assert_true(vim.fn.filereadable(src_fixture) == 1, "styled-test.docx fixture exists")

  local test_copy = test_dir .. "/test_styled.docx"
  vim.fn.system(string.format("cp %s %s", vim.fn.shellescape(src_fixture), vim.fn.shellescape(test_copy)))

  -- Open in Neovim
  vim.cmd("edit " .. vim.fn.fnameescape(test_copy))
  local bufnr = vim.api.nvim_get_current_buf()

  assert_equal(vim.bo[bufnr].filetype, "markdown", "filetype is markdown")
  assert_equal(vim.bo[bufnr].buftype, "acwrite", "buftype is acwrite")

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  assert_true(table.concat(lines, "\n"):find("python%-docx was here") ~= nil, "Original content loaded")

  -- Modify headings and text
  table.insert(lines, "# New Heading 1 Level")
  table.insert(lines, "## Subheading Level 2")
  table.insert(lines, "Testing style preservation with Pandoc reference-doc.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  -- Save
  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "modified is false after saving")

  -- Verify styles.xml is present and preserved
  local unzip_styles = vim.fn.system("unzip -l " .. vim.fn.shellescape(test_copy))
  assert_true(unzip_styles:find("word/styles%.xml") ~= nil, "styles.xml preserved in docx")

  -- Verify pandoc can read the modified docx cleanly
  local check_plain = vim.fn.system("pandoc " .. vim.fn.shellescape(test_copy) .. " -t plain")
  assert_true(check_plain:find("New Heading 1 Level") ~= nil, "New heading in saved document")
  assert_true(check_plain:find("Subheading Level 2") ~= nil, "Subheading in saved document")

  vim.cmd("bdelete!")
end

-- Test 6: Physical Word document with tables
tests["physical_docx_with_tables"] = function()
  local src_fixture = fixtures_dir .. "/tables.docx"
  assert_true(vim.fn.filereadable(src_fixture) == 1, "tables.docx fixture exists")

  local test_copy = test_dir .. "/test_tables.docx"
  vim.fn.system(string.format("cp %s %s", vim.fn.shellescape(src_fixture), vim.fn.shellescape(test_copy)))

  vim.cmd("edit " .. vim.fn.fnameescape(test_copy))
  local bufnr = vim.api.nvim_get_current_buf()

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  table.insert(lines, "Added paragraph between tables.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "saved tables docx")

  local check_plain = vim.fn.system("pandoc " .. vim.fn.shellescape(test_copy) .. " -t plain")
  assert_true(check_plain:find("Added paragraph between tables") ~= nil, "Table edits saved")

  vim.cmd("bdelete!")
end

-- Test 7: Special characters in filename (spaces, ampersands, hashes)
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

-- Test 8: Creating a brand new docx from scratch
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

-- Test 9: Save as new file (:w other_name.docx)
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

-- Test 10: Workspace cleanup on buffer wipeout
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

-- Test 11: Corrupted docx file handling
tests["corrupted_docx_handling"] = function()
  local corrupt_docx = test_dir .. "/corrupted.docx"
  vim.fn.writefile({ "This is not a zip or docx file at all!" }, corrupt_docx)

  local ok, err = pcall(function()
    vim.cmd("edit " .. vim.fn.fnameescape(corrupt_docx))
  end)
  assert_true(not ok, "Opening corrupted file correctly reports error")
  assert_true(tostring(err):find("[Pp]andoc") ~= nil, "Error mentions Pandoc: " .. tostring(err))
  pcall(vim.cmd, "bdelete!")
end

-- Test 12: Invalid pandoc binary path handling
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
