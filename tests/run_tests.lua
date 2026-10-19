vim.opt.rtp:prepend(".")
vim.cmd("runtime! plugin/nvim-docx.lua")

local zip = require("nvim-docx.zip")
local xml = require("nvim-docx.xml")
local ooxml = require("nvim-docx.ooxml")
local markdown = require("nvim-docx.markdown")
local converter = require("nvim-docx.converter")

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
local test_order = {}
local function add_test(name, fn)
  tests[name] = fn
  test_order[#test_order + 1] = name
end

local test_dir = vim.fn.tempname() .. "-nvim-docx-e2e"
local fixtures_dir = vim.fn.fnamemodify(debug.getinfo(1).source:sub(2), ":p:h") .. "/fixtures"

-- Pure-Lua generator for a valid 16x16 RGB PNG file (zero Python/external dependencies)
local function create_test_png(filepath)
  local function u32_be(n)
    local b1 = math.floor(n / 16777216) % 256
    local b2 = math.floor(n / 65536) % 256
    local b3 = math.floor(n / 256) % 256
    local b4 = n % 256
    return string.char(b1, b2, b3, b4)
  end

  local function adler32(data)
    local s1, s2 = 1, 0
    for i = 1, #data do
      s1 = (s1 + string.byte(data, i)) % 65521
      s2 = (s2 + s1) % 65521
    end
    return s2 * 65536 + s1
  end

  local function png_chunk(chunk_type, chunk_data)
    local crc = zip.crc32(chunk_type .. chunk_data)
    return u32_be(#chunk_data) .. chunk_type .. chunk_data .. u32_be(crc)
  end

  local sig = "\137PNG\r\n\26\n"
  -- IHDR: width=16, height=16, bit_depth=8, color_type=2 (RGB), compression=0, filter=0, interlace=0
  local ihdr_data = u32_be(16) .. u32_be(16) .. string.char(8, 2, 0, 0, 0)
  local row = "\0" .. string.rep("\31\74\140", 16)
  local raw = string.rep(row, 16)
  local zlib_stream = "\120\156" .. zip.deflate(raw) .. u32_be(adler32(raw))

  local png_bytes = sig
    .. png_chunk("IHDR", ihdr_data)
    .. png_chunk("IDAT", zlib_stream)
    .. png_chunk("IEND", "")

  local f = assert(io.open(filepath, "wb"))
  f:write(png_bytes)
  f:close()
end

local function read_docx_as_markdown(docx_path)
  local ws = vim.fn.tempname()
  vim.fn.mkdir(ws, "p")
  local lines, err = converter.docx_to_markdown_lines(docx_path, ws)
  vim.fn.delete(ws, "rf")
  assert_true(lines ~= nil, "read_docx_as_markdown succeeded: " .. tostring(err))
  return table.concat(lines, "\n")
end

local function setup_env()
  vim.fn.delete(test_dir, "rf")
  vim.fn.mkdir(test_dir, "p")
  vim.fn.mkdir(test_dir .. "/assets", "p")

  create_test_png(test_dir .. "/assets/logo.png")

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

  local ok, err = converter.markdown_to_docx(md_file, test_dir .. "/rich.docx", test_dir, nil)
  assert_true(ok, "Failed to generate rich.docx in setup_env: " .. tostring(err))
end

-- Test 1: Config and backward compatibility
add_test("config_defaults_and_setup", function()
  local config = require("nvim-docx.config")
  config.setup({
    debug = false,
    auto_cleanup = true,
  })
  assert_equal(config.options.debug, false, "debug")
  assert_equal(config.options.auto_cleanup, true, "auto_cleanup")
  assert_equal(config.options.preserve_styles, true, "preserve_styles default")
  assert_equal(config.options.wrap, "none", "wrap default")

  config.setup()
  assert_true(config.options ~= nil, "options exist after empty setup")
end)

-- Test 2: Checkhealth
add_test("checkhealth_runs_without_error", function()
  local ok, err = pcall(vim.cmd, "checkhealth nvim-docx")
  assert_true(ok, ":checkhealth nvim-docx succeeded: " .. tostring(err))
  local report = table.concat(vim.api.nvim_buf_get_lines(0, 0, -1, false), "\n")
  assert_true(report:find("engine active", 1, true) ~= nil, "health report mentions the engine: " .. report)
  assert_true(report:find("ERROR", 1, true) == nil, "health report has no errors: " .. report)
  vim.cmd("bwipeout!")
end)

-- Test 3: Open rich docx, verify markdown buffer, edit, save, verify docx
add_test("open_edit_save_rich_docx", function()
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
  assert_true(text:find("%*%*bold%*%*") ~= nil, "Contains bold formatting")
  assert_true(text:find("%*italic%*") ~= nil, "Contains italic formatting")
  assert_true(text:find("https://github%.com/DaanHessen/nvim%-docx") ~= nil, "Contains hyperlink")
  assert_true(text:find("Item Alpha") ~= nil, "Contains list item")
  assert_true(text:find("  %- Subitem Beta%.1") ~= nil, "Contains nested list item")
  assert_true(text:find("Col 1") ~= nil, "Contains table column")
  assert_true(text:find("logo%.png") ~= nil or text:find("media") ~= nil, "Contains image reference")

  table.insert(lines, "Added by automated test verification.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)
  assert_true(vim.bo[bufnr].modified, "buffer is marked modified")

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "buffer marked unmodified after write")

  local arch = assert(zip.read_zip(rich_docx))
  local has_media = false
  for _, name in ipairs(zip.names(arch)) do
    if name:find("^word/media/") then
      has_media = true
      break
    end
  end
  assert_true(has_media, "DOCX still contains media archive")

  local check_text = read_docx_as_markdown(rich_docx)
  assert_true(check_text:find("Added by automated test verification") ~= nil, "Saved docx contains edited text")

  vim.cmd("bdelete!")
end)

-- Test 4: Physical Microsoft Word document with embedded images
add_test("physical_docx_with_embedded_images", function()
  local src_fixture = fixtures_dir .. "/having-images.docx"
  assert_true(vim.fn.filereadable(src_fixture) == 1, "having-images.docx fixture exists")

  local test_copy = test_dir .. "/test_having_images.docx"
  local init_arch = assert(zip.read_zip(src_fixture))
  assert_true(init_arch.files["word/media/image1.png"] ~= nil, "Fixture contains image1.png")
  assert_true(init_arch.files["word/media/image2.png"] ~= nil, "Fixture contains image2.png")
  assert(zip.write_zip(test_copy, init_arch))

  -- Open in Neovim
  vim.cmd("edit " .. vim.fn.fnameescape(test_copy))
  local bufnr = vim.api.nvim_get_current_buf()

  assert_equal(vim.bo[bufnr].filetype, "markdown", "filetype is markdown")
  assert_equal(vim.bo[bufnr].buftype, "acwrite", "buftype is acwrite")

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local text = table.concat(lines, "\n")
  assert_true(text:find("media/image1%.png") ~= nil, "Extracted media referenced in markdown buffer: " .. text)

  -- Edit buffer: add custom heading and text
  table.insert(lines, 1, "# Customized Heading in Image Document")
  table.insert(lines, 2, "Verified that images remain intact after editing.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  -- Save
  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "modified is false after saving")

  -- Verify images and headers are STILL present in the saved DOCX zip archive
  local saved_arch = assert(zip.read_zip(test_copy))
  assert_true(saved_arch.files["word/media/image1.png"] ~= nil, "image1.png preserved in saved docx")
  assert_true(saved_arch.files["word/media/image2.png"] ~= nil, "image2.png preserved in saved docx")
  assert_true(saved_arch.files["word/header1.xml"] ~= nil, "header1.xml preserved in saved docx")

  -- Re-open in Neovim to verify roundtrip fidelity
  vim.cmd("edit!")
  local reloaded_lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local reloaded_text = table.concat(reloaded_lines, "\n")
  assert_true(reloaded_text:find("Customized Heading in Image Document") ~= nil, "Reloaded buffer has edited heading")
  assert_true(reloaded_text:find("media/image1%.png") ~= nil, "Reloaded buffer retains image references")

  vim.cmd("bdelete!")
end)

-- Test 5: Physical Microsoft Word document with custom styles and fonts
add_test("physical_docx_styled_document", function()
  local src_fixture = fixtures_dir .. "/styled-test.docx"
  assert_true(vim.fn.filereadable(src_fixture) == 1, "styled-test.docx fixture exists")

  local test_copy = test_dir .. "/test_styled.docx"
  local init_arch = assert(zip.read_zip(src_fixture))
  assert(zip.write_zip(test_copy, init_arch))

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
  table.insert(lines, "Testing style preservation with native OOXML engine.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  -- Save
  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "modified is false after saving")

  -- Verify styles.xml and customXml are present and preserved
  local saved_arch = assert(zip.read_zip(test_copy))
  assert_true(saved_arch.files["word/styles.xml"] ~= nil, "styles.xml preserved in docx")
  assert_true(saved_arch.files["customXml/item1.xml"] ~= nil, "customXml preserved in docx")

  local check_plain = read_docx_as_markdown(test_copy)
  assert_true(check_plain:find("New Heading 1 Level") ~= nil, "New heading in saved document")
  assert_true(check_plain:find("Subheading Level 2") ~= nil, "Subheading in saved document")

  vim.cmd("bdelete!")
end)

-- Test 6: Physical Word document with tables
add_test("physical_docx_with_tables", function()
  local src_fixture = fixtures_dir .. "/tables.docx"
  assert_true(vim.fn.filereadable(src_fixture) == 1, "tables.docx fixture exists")

  local test_copy = test_dir .. "/test_tables.docx"
  local init_arch = assert(zip.read_zip(src_fixture))
  assert(zip.write_zip(test_copy, init_arch))

  vim.cmd("edit " .. vim.fn.fnameescape(test_copy))
  local bufnr = vim.api.nvim_get_current_buf()

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local initial_md = table.concat(lines, "\n")
  assert_true(initial_md:find("| T2 |") ~= nil, "Tables rendered as clean Markdown pipe tables")

  table.insert(lines, "Added paragraph between tables.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "saved tables docx")

  local check_md = read_docx_as_markdown(test_copy)
  assert_true(check_md:find("Added paragraph between tables") ~= nil, "Table edits saved")
  assert_true(check_md:find("| T2 |") ~= nil, "Table content preserved after save")

  vim.cmd("bdelete!")
end)

-- Test 7: Special characters in filename (spaces, ampersands, hashes)
add_test("filename_with_spaces_and_symbols", function()
  local tricky_name = test_dir .. "/Report - 2026 & Test (Special) #1.docx"
  local init_arch = assert(zip.read_zip(test_dir .. "/rich.docx"))
  assert(zip.write_zip(tricky_name, init_arch))

  vim.cmd("edit " .. vim.fn.fnameescape(tricky_name))
  local bufnr = vim.api.nvim_get_current_buf()

  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  table.insert(lines, "Tricky filename test line.")
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)

  vim.cmd("write")
  assert_equal(vim.bo[bufnr].modified, false, "saved tricky filename")

  local check_text = read_docx_as_markdown(tricky_name)
  assert_true(check_text:find("Tricky filename test line") ~= nil, "Content persisted in tricky filename")

  vim.cmd("bdelete!")
end)

-- Test 8: Creating a brand new docx from scratch
add_test("create_new_docx_from_scratch", function()
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

  local check_out = read_docx_as_markdown(new_docx)
  assert_true(check_out:find("# Brand New Document") ~= nil, "New DOCX contains markdown heading")
  assert_true(check_out:find("%- Item 1") ~= nil, "New DOCX contains bullet list")

  vim.cmd("bdelete!")
end)

-- Test 9: Save as new file (:w other_name.docx)
add_test("save_as_new_filename", function()
  local orig_docx = test_dir .. "/rich.docx"
  local copy_docx = test_dir .. "/copy_via_saveas.docx"
  vim.fn.delete(copy_docx)

  vim.cmd("edit " .. vim.fn.fnameescape(orig_docx))
  vim.cmd("write " .. vim.fn.fnameescape(copy_docx))
  assert_true(vim.fn.filereadable(copy_docx) == 1, "copy_via_saveas.docx was created")

  local check_out = read_docx_as_markdown(copy_docx)
  assert_true(check_out:find("Test Document Header") ~= nil, "Copy contains original content")

  vim.cmd("bdelete!")
end)

-- Test 10: Workspace cleanup on buffer wipeout
add_test("workspace_lifecycle_cleanup", function()
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
end)

-- Test 11: Corrupted docx file handling
add_test("corrupted_docx_handling", function()
  local corrupt_docx = test_dir .. "/corrupted.docx"
  vim.fn.writefile({ "This is not a zip or docx file at all!" }, corrupt_docx)

  local ok, err = pcall(function()
    vim.cmd("edit " .. vim.fn.fnameescape(corrupt_docx))
  end)
  assert_true(not ok, "Opening corrupted file correctly reports error")
  assert_true(tostring(err):find("ZIP") ~= nil or tostring(err):find("DOCX") ~= nil, "Error mentions invalid ZIP/DOCX: " .. tostring(err))
  pcall(vim.cmd, "bdelete!")
end)

-- Test 12: Rich Markdown constructs (code blocks, blockquotes, strikethrough, sup/sub, tasks, footnotes, hr)
add_test("advanced_markdown_constructs_roundtrip", function()
  local adv_docx = test_dir .. "/advanced.docx"
  vim.fn.delete(adv_docx)

  vim.cmd("edit " .. vim.fn.fnameescape(adv_docx))
  local bufnr = vim.api.nvim_get_current_buf()

  local adv_lines = {
    "# Advanced Features",
    "",
    "Text with **bold**, *italic*, ***both***, ~~strikethrough~~, `inline_code`, H~2~O, X^2^, and footnote[^note1].",
    "",
    "> Blockquote wisdom line.",
    "",
    "- [x] Finished task",
    "- [ ] Pending task",
    "",
    "1. Ordered first",
    "2. Ordered second",
    "",
    "```lua",
    "local function hello()",
    "  return 42",
    "end",
    "```",
    "",
    "---",
    "",
    "[^note1]: Detailed footnote explanation.",
  }

  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, adv_lines)
  vim.cmd("write")
  vim.cmd("edit!")

  local reloaded = table.concat(vim.api.nvim_buf_get_lines(bufnr, 0, -1, false), "\n")
  assert_true(reloaded:find("~~strikethrough~~") ~= nil, "Strikethrough preserved")
  assert_true(reloaded:find("`inline_code`") ~= nil, "Inline code preserved")
  assert_true(reloaded:find("H~2~O") ~= nil, "Subscript preserved")
  assert_true(reloaded:find("X%^2%^") ~= nil, "Superscript preserved")
  assert_true(reloaded:find("> Blockquote wisdom line%.") ~= nil, "Blockquote preserved")
  assert_true(reloaded:find("%- %[x%] Finished task") ~= nil, "Checked task preserved")
  assert_true(reloaded:find("%- %[ %] Pending task") ~= nil, "Unchecked task preserved")
  assert_true(reloaded:find("1%. Ordered first") ~= nil, "Ordered list 1 preserved")
  assert_true(reloaded:find("2%. Ordered second") ~= nil, "Ordered list 2 preserved")
  assert_true(reloaded:find("local function hello%(%)") ~= nil, "Code block preserved")
  assert_true(reloaded:find("%-%-%-") ~= nil, "Horizontal rule preserved")
  assert_true(reloaded:find("%[%^1%]: Detailed footnote explanation%.") ~= nil, "Footnote preserved")

  vim.cmd("bdelete!")
end)

-- Test 13: Track changes handling ("accept", "reject", "all")
add_test("track_changes_modes", function()
  local base_arch = ooxml.ast_to_docx(markdown.markdown_to_ast({ "Base" }), test_dir, nil, {})
  base_arch.files["word/document.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:body><w:p>',
    '<w:r><w:t xml:space="preserve">Keep </w:t></w:r>',
    '<w:ins w:id="1" w:author="Test"><w:r><w:t xml:space="preserve">Inserted</w:t></w:r></w:ins>',
    '<w:del w:id="2" w:author="Test"><w:r><w:delText xml:space="preserve">Deleted</w:delText></w:r></w:del>',
    '</w:p></w:body></w:document>',
  })

  local ast_accept = ooxml.docx_to_ast(base_arch, nil, { track_changes = "accept" })
  local md_accept = markdown.ast_to_markdown(ast_accept)
  assert_equal(md_accept, "Keep Inserted", "track_changes=accept keeps insertions and drops deletions")

  local ast_reject = ooxml.docx_to_ast(base_arch, nil, { track_changes = "reject" })
  local md_reject = markdown.ast_to_markdown(ast_reject)
  assert_equal(md_reject, "Keep ~~Deleted~~", "track_changes=reject keeps deletions and drops insertions")

  local ast_all = ooxml.docx_to_ast(base_arch, nil, { track_changes = "all" })
  local md_all = markdown.ast_to_markdown(ast_all)
  assert_equal(md_all, "Keep Inserted~~Deleted~~", "track_changes=all includes both")
end)

-- Test 14: User commands (:DocxInfo, :DocxExportMarkdown, :DocxFromMarkdown)
add_test("user_commands_info_export_from_markdown", function()
  local rich_docx = test_dir .. "/rich.docx"
  vim.cmd("edit " .. vim.fn.fnameescape(rich_docx))

  local stats = require("nvim-docx").docx_info()
  assert_true(stats ~= nil, "docx_info returned stats table")
  assert_true(stats.headings >= 2, "stats counted headings")
  assert_true(stats.tables >= 1, "stats counted tables")
  assert_true(stats.images >= 1, "stats counted images")

  local exported_md = test_dir .. "/exported_output.md"
  vim.cmd("DocxExportMarkdown " .. vim.fn.fnameescape(exported_md))
  assert_true(vim.fn.filereadable(exported_md) == 1, "DocxExportMarkdown created .md file")
  vim.cmd("bdelete!")

  -- Open the exported .md file and compile it to a new .docx via :DocxFromMarkdown
  local compiled_docx = test_dir .. "/compiled_from_md.docx"
  vim.cmd("edit " .. vim.fn.fnameescape(exported_md))
  vim.cmd("DocxFromMarkdown " .. vim.fn.fnameescape(compiled_docx))
  assert_true(vim.fn.filereadable(compiled_docx) == 1, "DocxFromMarkdown created .docx file")
  vim.cmd("bdelete!")
end)

-- Test 15: Pure-Lua ZIP and XML unit verification
add_test("zip_and_xml_unit_verification", function()
  local text = string.rep("Hello OOXML World! <tag attr=\"val\">&amp;</tag> ", 50)
  local compressed = zip.deflate(text)
  assert_true(#compressed < #text, "DEFLATE compressed repetitive string")
  local decompressed = zip.inflate(compressed)
  assert_equal(decompressed, text, "INFLATE restored exact original bytes")

  local sample_xml = '<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE root [<!ENTITY foo "bar">]><w:p><w:r><w:t>Caf&#233; &amp; &#x20AC;</w:t></w:r></w:p>'
  local root = assert(xml.parse(sample_xml))
  local t_node = xml.find_all(root, "t")[1]
  assert_equal(xml.get_direct_text(t_node), "Café & €", "XML numeric and named entities decoded to UTF-8")
end)

-- Test 16: Nested bold/italic combinations and 4-space indented nested lists
add_test("nested_emphasis_and_4space_lists", function()
  local md_input = {
    "Paragraph with *italic and **bold** inside* and **bold with *italic* inside** and *italic **both***.",
    "",
    "1. First outer",
    "    - Nested bullet with 4 spaces",
    "        - Deeply nested with 8 spaces",
    "2. Second outer continues at 2",
  }
  local ast = markdown.markdown_to_ast(md_input)
  local arch = ooxml.ast_to_docx(ast, test_dir, nil, {})
  local rt_ast = ooxml.docx_to_ast(arch, test_dir, {})
  local rt_md = markdown.ast_to_markdown(rt_ast)

  assert_true(rt_md:find("%*italic and %*%*bold%*%* inside%*") or rt_md:find("%*%*bold%*%*"), "Nested bold inside italic preserved: " .. rt_md)
  assert_true(rt_md:find("  %- Nested bullet with 4 spaces") ~= nil, "4-space indent mapped to level 1 without skipping: " .. rt_md)
  assert_true(rt_md:find("    %- Deeply nested with 8 spaces") ~= nil, "8-space indent mapped to level 2: " .. rt_md)
  assert_true(rt_md:find("2%. Second outer continues at 2") ~= nil, "Outer ordered list numbering preserved after nested bullet list: " .. rt_md)
end)

-- Test 17: Inline math ($...$), display math blocks ($$...$$), and structured Office Math (<m:oMath>)
add_test("office_math_equations_roundtrip", function()
  local md_input = {
    "Einstein wrote $E = mc^2$ in inline math.",
    "",
    "$$",
    "\\int_0^\\infty e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}",
    "$$",
  }
  local ast = markdown.markdown_to_ast(md_input)
  local arch = ooxml.ast_to_docx(ast, test_dir, nil, {})
  local doc_xml = arch.files["word/document.xml"]
  assert_true(doc_xml:find("<m:nary>") ~= nil, "Generated structured <m:nary> integral element")
  assert_true(doc_xml:find("<m:f>") ~= nil, "Generated structured <m:f> fraction element")
  assert_true(doc_xml:find("<m:rad>") ~= nil, "Generated structured <m:rad> radical element")
  assert_true(doc_xml:find("<m:sSup>") ~= nil, "Generated structured <m:sSup> superscript element")

  local rt_ast = ooxml.docx_to_ast(arch, test_dir, {})
  local rt_md = markdown.ast_to_markdown(rt_ast)

  assert_true(rt_md:find("%$E = mc%^{2}%$") ~= nil, "Inline math preserved in roundtrip: " .. rt_md)
  assert_true(rt_md:find("%$%$\n\\int_{0}%^{\\infty} e%^{%-x%^{2}} dx = \\frac{\\sqrt{\\pi}}{2}") ~= nil, "Display math block preserved in roundtrip: " .. rt_md)

  -- Also test structured OOXML math elements (<m:f>, <m:sSup>, <m:rad>)
  arch.files["word/document.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">',
    '<w:body><w:p>',
    '<m:oMath>',
      '<m:f><m:num><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:num>',
      '<m:den><m:rad><m:deg/><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad></m:den></m:f>',
    '</m:oMath>',
    '</w:p></w:body></w:document>',
  })
  local struct_ast = ooxml.docx_to_ast(arch, nil, {})
  local struct_md = markdown.ast_to_markdown(struct_ast)
  assert_equal(struct_md, "$\\frac{x^{2}}{\\sqrt{y}}$", "Structured Office Math converted to LaTeX math")
end)

-- Test 18: Complex Word HYPERLINK fields (w:fldChar + w:instrText) and Word comments
add_test("complex_fields_and_word_comments", function()
  local arch = ooxml.ast_to_docx(markdown.markdown_to_ast({ "Placeholder" }), test_dir, nil, {})
  arch.files["word/comments.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:comment w:id="0" w:author="Alice">',
      '<w:p><w:r><w:t>Please review this link.</w:t></w:r></w:p>',
    '</w:comment>',
    '</w:comments>',
  })
  arch.files["word/document.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:body><w:p>',
    '<w:r><w:t xml:space="preserve">Visit </w:t></w:r>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://neovim.io" \\o "Neovim" </w:instrText></w:r>',
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    '<w:r><w:t>Neovim Website</w:t></w:r>',
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    '<w:r><w:commentReference w:id="0"/></w:r>',
    '</w:p></w:body></w:document>',
  })

  local ast_all = ooxml.docx_to_ast(arch, nil, { track_changes = "all" })
  local md_all = markdown.ast_to_markdown(ast_all)
  assert_true(md_all:find("%[Neovim Website%]%(https://neovim%.io%)") ~= nil, "Complex HYPERLINK field extracted: " .. md_all)
  assert_true(md_all:find("%[%^1%]") ~= nil, "Comment reference extracted in track_changes=all mode: " .. md_all)
  assert_true(md_all:find("%*%*%[Alice%]:%*%* Please review this link%.") ~= nil, "Comment body and author extracted: " .. md_all)
end)

-- Test 19: Workspace leak prevention on :edit! and :bdelete!
add_test("no_workspace_leaks_on_reload_or_bdelete", function()
  local docx = test_dir .. "/rich.docx"
  vim.cmd("edit " .. vim.fn.fnameescape(docx))
  local bufnr = vim.api.nvim_get_current_buf()
  local ws1 = vim.b[bufnr].nvim_docx_state.workspace
  assert_true(vim.fn.isdirectory(ws1) == 1, "Initial workspace exists")

  -- Reload buffer with :edit!
  vim.cmd("edit!")
  local ws2 = vim.b[bufnr].nvim_docx_state.workspace
  assert_true(ws1 ~= ws2, "Reload created a fresh workspace")
  assert_equal(vim.fn.isdirectory(ws1), 0, "Previous workspace cleaned up on :edit!")
  assert_true(vim.fn.isdirectory(ws2) == 1, "New workspace exists after :edit!")

  -- Close buffer with :bdelete! (not bwipeout)
  vim.cmd("bdelete!")
  assert_equal(vim.fn.isdirectory(ws2), 0, "Workspace cleaned up on :bdelete!")
end)

-- Test 20: Real-world OOXML quirks (numId=0 suppression, startOverride restart, verbatim sup/sub, VML & external images)
add_test("real_world_ooxml_quirks_roundtrip", function()
  -- 1. Verbatim inside superscript/subscript (`m`^`2`^ and `m`~`2`~)
  local v_ast = markdown.markdown_to_ast({ "`m`^`2`^ and `m`~`2`~" })
  local v_md = markdown.ast_to_markdown(v_ast)
  assert_equal(v_md, "`m`^`2`^ and `m`~`2`~", "Verbatim superscript/subscript idempotency")

  -- 2. numId=0 list suppression, startOverride list restart, VML image dimensions, and external image roundtrip
  local arch = ooxml.ast_to_docx(markdown.markdown_to_ast({ "Init" }), test_dir, nil, {})
  arch.files["word/numbering.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:abstractNum w:abstractNumId="0">',
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="ListNumber"/></w:lvl>',
    '</w:abstractNum>',
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
    '<w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>',
    '</w:numbering>',
  })
  arch.files["word/styles.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:style w:type="paragraph" w:styleId="ListNumber">',
      '<w:name w:val="List Number"/>',
      '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>',
    '</w:style>',
    '</w:styles>',
  })
  arch.files["word/_rels/document.xml.rels"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="external-pic.png" TargetMode="External"/>',
    '</Relationships>',
  })
  arch.files["word/document.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
    'xmlns:v="urn:schemas-microsoft-com:vml">',
    '<w:body>',
      '<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Suppressed list item   </w:t></w:r></w:p>',
      '<w:p><w:r><w:tab/></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Restarted 1</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>Continued 2</w:t></w:r></w:p>',
      '<w:p><w:r><w:pict><v:rect style="width:72pt;height:36pt"><v:imagedata r:id="rId10"/></v:rect></w:pict></w:r></w:p>',
      '<w:tbl>',
        '<w:tr><w:tc><w:p><w:r><w:t> H1 </w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t> H2 </w:t></w:r></w:p></w:tc></w:tr>',
        '<w:tr><w:tc><w:p><w:r><w:t>Aliquam erat volutpat.  </w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>  Lorem  </w:t></w:r></w:p></w:tc></w:tr>',
      '</w:tbl>',
    '</w:body></w:document>',
  })

  local ast1 = ooxml.docx_to_ast(arch, test_dir, {})
  local md1 = markdown.ast_to_markdown(ast1)
  assert_true(md1:find("1%. First\n\nSuppressed list item\n\n1%. Restarted 1\n2%. Continued 2") ~= nil, "numId=0 suppression, tab-only paragraph removal, and startOverride restart: " .. md1)
  assert_true(md1:find('!%[%]%(external%-pic%.png%){width="1in" height="0%.5in"}') ~= nil, "VML style dimensions and external image extracted: " .. md1)
  assert_true(md1:find("| Aliquam erat volutpat%. | Lorem |") ~= nil, "Table cell leading/trailing spaces trimmed for idempotency: " .. md1)

  -- Round-trip back to DOCX and verify external image + dimensions + table cells survive
  local ast2 = markdown.markdown_to_ast(vim.split(md1, "\n", { plain = true }))
  local arch2 = ooxml.ast_to_docx(ast2, test_dir, arch, { preserve_styles = true })
  local ast3 = ooxml.docx_to_ast(arch2, test_dir, {})
  local md3 = markdown.ast_to_markdown(ast3)
  assert_equal(md3, md1, "Full round-trip preserves external image dimensions, list numbering, and trimmed table cells")

  -- Legacy binary .doc (OLE2 Compound Document) detection
  local ole2_bytes = "\208\207\017\224\161\177\026\225" .. string.rep("\0", 32)
  local ole_arch, ole_err = zip.read_zip_bytes(ole2_bytes)
  assert_true(ole_arch == nil, "OLE2 binary .doc rejected")
  assert_true((ole_err or ""):find("OLE2") ~= nil, "OLE2 error mentions legacy .doc format: " .. tostring(ole_err))
end)

-- Test 21: AlternateContent, TextBoxes (txbxContent), SmartArt diagrams, w:sym, Underline, and Charts (<c:chart>)
add_test("advanced_ooxml_containers_and_charts", function()
  local md_u = markdown.markdown_to_ast({ "Testing <u>underlined</u> and [also underlined]{.underline} text." })
  local arch_u = ooxml.ast_to_docx(md_u, test_dir, nil, {})
  assert_true(arch_u.files["word/document.xml"]:find('<w:u w:val="single"/>') ~= nil, "<u> and {.underline} emit <w:u>")

  local arch = ooxml.ast_to_docx(markdown.markdown_to_ast({ "Base" }), test_dir, nil, {})
  arch.files["word/charts/chart1.xml"] = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>'
  arch.files["word/diagrams/data1.xml"] = '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt><dgm:t><a:p><a:r><a:t>SmartArt Node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>'
  arch.files["word/_rels/document.xml.rels"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>',
    '<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>',
    '</Relationships>',
  })
  arch.files["word/document.xml"] = table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ',
    'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<w:body>',
      '<w:p><w:r>',
        '<mc:AlternateContent>',
          '<mc:Choice Requires="wps"><w:drawing><w:txbxContent><w:p><w:r><w:t>TextBox Caption</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice>',
          '<mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>TextBox Caption</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback>',
        '</mc:AlternateContent>',
      '</w:r></w:p>',
      '<w:p><w:r><w:sym w:font="Symbol" w:char="F0B1"/><w:softHyphen/><w:t>5</w:t></w:r></w:p>',
      '<w:p><w:r><w:drawing><a:graphic><a:graphicData><dgm:relIds r:dm="rId21"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>',
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="3657600" cy="1828800"/><wp:docPr id="1" name="Sales Chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId20"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    '</w:body></w:document>',
  })

  local ast1 = ooxml.docx_to_ast(arch, test_dir, {})
  local md1 = markdown.ast_to_markdown(ast1)
  local _, count_tb = md1:gsub("TextBox Caption", "")
  assert_equal(count_tb, 1, "AlternateContent extracted Choice without duplicating Fallback: " .. md1)
  assert_true(md1:find("\194\177\194\1735") ~= nil, "w:sym (F0B1 -> ±) and w:softHyphen decoded: " .. md1)
  assert_true(md1:find("SmartArt Node") ~= nil, "SmartArt diagram text extracted: " .. md1)
  assert_true(md1:find('!%[Sales Chart%]%(charts/chart1%.xml%){width="4in" height="2in"}') ~= nil, "Embedded chart extracted with dimensions: " .. md1)

  -- Round-trip chart back to DOCX and verify <c:chart> and chart1.xml survive
  local ast2 = markdown.markdown_to_ast(vim.split(md1, "\n", { plain = true }))
  local arch2 = ooxml.ast_to_docx(ast2, test_dir, arch, { preserve_styles = true })
  assert_true(arch2.files["word/charts/chart1.xml"] ~= nil, "chart1.xml preserved in archive")
  assert_true(arch2.files["word/document.xml"]:find("<c:chart ") ~= nil, "<c:chart> element emitted in document.xml")
  local ast3 = ooxml.docx_to_ast(arch2, test_dir, {})
  local md3 = markdown.ast_to_markdown(ast3)
  assert_equal(md3, md1, "Full round-trip preserves chart, SmartArt text, TextBox text, and symbols")
end)

add_test("lossless_save_keeps_direct_formatting_of_untouched_and_edited_blocks", function()
  local docx = test_dir .. "/lossless.docx"
  local ws = vim.fn.tempname()
  vim.fn.mkdir(ws, "p")
  assert_true(converter.markdown_lines_to_docx({
    "::: {font=\"Georgia\" size=15}",
    "",
    "First paragraph in Georgia.",
    "",
    "Second paragraph in Georgia.",
    "",
    ":::",
    "",
    "Plain [big]{size=20} word.",
  }, docx, ws, nil, {}), "fixture written")

  vim.cmd("edit " .. vim.fn.fnameescape(docx))
  local bufnr = vim.api.nvim_get_current_buf()
  local lines = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  local text = table.concat(lines, "\n")
  assert_true(text:find("[big]{size=20}", 1, true) ~= nil, "differing run size is shown: " .. text)
  assert_true(text:find("font=\"Georgia\"", 1, true) == nil, "paragraph's own font is not shown: " .. text)
  for i, l in ipairs(lines) do
    if l:find("Second paragraph", 1, true) then
      lines[i] = "Second paragraph, edited in Neovim."
    end
  end
  vim.api.nvim_buf_set_lines(bufnr, 0, -1, false, lines)
  vim.cmd("write")
  vim.cmd("bwipeout!")

  local arch = assert(zip.read_zip(docx))
  local doc = arch.files["word/document.xml"]
  assert_true(doc:find("edited in Neovim", 1, true) ~= nil, "edit saved")
  assert_true(select(2, doc:gsub('w:ascii="Georgia"', "")) >= 2, "Georgia kept on untouched and edited paragraphs")
  assert_true(doc:find('w:sz w:val="40"', 1, true) ~= nil, "20 pt run kept")
  assert_true(select(2, doc:gsub('w:sz w:val="30"', "")) >= 2, "15 pt kept on both paragraphs")
end)

add_test("attribute_spans_sections_and_page_breaks_roundtrip", function()
  local src = {
    "::: {style=\"Quote\"}",
    "",
    "A quote paragraph.",
    "",
    ":::",
    "",
    "Text with [red 14 pt]{size=14 color=#FF0000} and [marked]{highlight=yellow}.",
    "",
    "\\newpage",
    "",
    "Literal *stars*, a\\_b, \\$5 and ***bold italic*** stay intact: \\*not emphasis\\*.",
  }
  local out = test_dir .. "/attrs.docx"
  local ws = vim.fn.tempname()
  vim.fn.mkdir(ws, "p")
  assert_true(converter.markdown_lines_to_docx(src, out, ws, nil, {}), "attrs doc written")
  local arch = assert(zip.read_zip(out))
  local doc = arch.files["word/document.xml"]
  assert_true(doc:find('<w:pStyle w:val="Quote"/>', 1, true) ~= nil, "Quote style applied")
  assert_true(doc:find('<w:color w:val="FF0000"/>', 1, true) ~= nil, "color applied")
  assert_true(doc:find('<w:highlight w:val="yellow"/>', 1, true) ~= nil, "highlight applied")
  assert_true(doc:find('<w:br w:type="page"/>', 1, true) ~= nil, "page break written")
  local back = table.concat(converter.docx_to_markdown_lines(out, nil), "\n")
  assert_true(back:find("[red 14 pt]{size=14 color=#FF0000}", 1, true) ~= nil, "span read back: " .. back)
  assert_true(back:find("\\newpage", 1, true) ~= nil, "page break read back")
  assert_true(back:find("Literal *stars*, a_b, \\$5 and ***bold italic*** stay intact: \\*not emphasis\\*.", 1, true) ~= nil,
    "escaping round-trips: " .. back)
end)

add_test("live_preview_renders_html", function()
  local preview = require("nvim-docx.preview")
  local html = preview.render_html({
    "# Title",
    "",
    "Some **bold** and [red]{color=red}.",
    "",
    "- a",
    "- b",
    "1. one",
    "",
    "| A | B |",
    "| :--- | ---: |",
    "| 1 | 2 |",
  })
  assert_true(html:find("<h1>Title</h1>", 1, true) ~= nil, "heading: " .. html)
  assert_true(html:find("<strong>bold</strong>", 1, true) ~= nil, "bold")
  assert_true(html:find("color:#FF0000", 1, true) ~= nil, "span color")
  assert_true(html:find("</ul>\n<ol>", 1, true) ~= nil, "bullet then numbered list: " .. html)
  assert_true(html:find('<td style="text-align:right">2</td>', 1, true) ~= nil, "table alignment")
end)

add_test("zip_plugin_does_not_claim_word_files", function()
  local docx = test_dir .. "/rich.docx"
  vim.cmd("edit " .. vim.fn.fnameescape(docx))
  assert_equal(vim.bo.readonly, false, "buffer is not readonly")
  assert_equal(vim.bo.buftype, "acwrite", "buffer is acwrite")
  assert_equal(vim.bo.filetype, "markdown", "filetype markdown")
  vim.cmd("bwipeout!")
end)

add_test("18. Undo safety on open, .DOCX uppercase autocmd, numbering xmlns preservation, non-default officeDocument rels, and VP8L/TIFF/SVG dimensions", function()
  local image = require("nvim-docx.image")
  local ooxml = require("nvim-docx.ooxml")
  local zip = require("nvim-docx.zip")
  local docx = require("nvim-docx")

  -- 1. VP8L lossless WebP dimensions (width=320, height=240)
  -- bits = (w - 1) | ((h - 1) << 14) = 319 + 239 * 16384 = 3916095 = 0x003BC13F
  local vp8l_bytes = "RIFF\24\0\0\0WEBPVP8L\12\0\0\0\47\63\193\59\0\0\0\0\0"
  local vw, vh = image.get_pixel_dimensions(vp8l_bytes)
  assert_equal(vw, 320, "VP8L width")
  assert_equal(vh, 240, "VP8L height")

  -- 2. Little-endian TIFF (width=640, height=480)
  -- II\42\0, offset=8, 2 entries: tag 0x0100 SHORT(3) count=1 val=640 (0x0280), tag 0x0101 SHORT(3) count=1 val=480 (0x01E0)
  local tiff_le = "II\42\0\8\0\0\0\2\0"
    .. "\0\1\3\0\1\0\0\0\128\2\0\0"
    .. "\1\1\3\0\1\0\0\0\224\1\0\0"
  local tw, th = image.get_pixel_dimensions(tiff_le)
  assert_equal(tw, 640, "TIFF LE width")
  assert_equal(th, 480, "TIFF LE height")

  -- 3. SVG dimensions from width/height and from viewBox
  local svg1 = '<svg xmlns="http://www.w3.org/2000/svg" width="200px" height="100px"><rect/></svg>'
  local sw1, sh1 = image.get_pixel_dimensions(svg1)
  assert_equal(sw1, 200, "SVG width attribute")
  assert_equal(sh1, 100, "SVG height attribute")

  local svg2 = '<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 500 250"><rect/></svg>'
  local sw2, sh2 = image.get_pixel_dimensions(svg2)
  assert_equal(sw2, 500, "SVG viewBox width fallback")
  assert_equal(sh2, 250, "SVG viewBox height fallback")

  -- 4. Preserving extra xmlns:* attributes on numbering.xml (e.g., xmlns:w15) & non-default officeDocument path
  local base_ast = markdown.markdown_to_ast({ "- Item 1" })
  local arch = ooxml.ast_to_docx(base_ast, test_dir, nil, {})
  -- Inject w15:restartNumberingAfterBreak on numbering.xml and relocate document.xml to document2.xml
  arch.files["word/numbering.xml"] = arch.files["word/numbering.xml"]
    :gsub('<w:numbering ', '<w:numbering xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ')
    :gsub('<w:abstractNum w:abstractNumId="0">', '<w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">')
  arch.files["word/document2.xml"] = arch.files["word/document.xml"]
  arch.files["word/document.xml"] = nil
  arch.files["_rels/.rels"] = arch.files["_rels/.rels"]:gsub("word/document%.xml", "word/document2.xml")

  local arch_rt = ooxml.ast_to_docx(base_ast, test_dir, arch, { preserve_styles = true })
  assert_true(arch_rt.files["word/numbering.xml"]:find('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"') ~= nil, "numbering.xml preserves xmlns:w15 declaration")
  assert_true(arch_rt.files["_rels/.rels"]:find('Target="word/document.xml"') ~= nil, "_rels/.rels updated to point to word/document.xml")

  -- 5. Uppercase .DOCX autocmd + pressing `u` immediately after opening does NOT wipe buffer
  local upper_file = test_dir .. "/UPPERCASE_TEST.DOCX"
  zip.write_zip(upper_file, arch_rt)
  vim.cmd("edit " .. vim.fn.fnameescape(upper_file))
  local bufnr = vim.api.nvim_get_current_buf()
  local lines_before = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  assert_true(table.concat(lines_before, "\n"):find("Item 1") ~= nil, "Uppercase .DOCX loaded via autocmd")

  -- Execute normal! u (undo) right after opening
  vim.cmd("silent! normal! u")
  local lines_after_undo = vim.api.nvim_buf_get_lines(bufnr, 0, -1, false)
  assert_equal(table.concat(lines_after_undo, "\n"), table.concat(lines_before, "\n"), "Pressing u after opening .docx does not wipe initial buffer content")

  -- Also verify :DocxInfo runs without error on open buffer
  docx.docx_info(bufnr)
end)

-- Main runner
local ok_setup, setup_err = pcall(setup_env)
if not ok_setup then
  print("FATAL: setup_env failed: " .. tostring(setup_err))
  vim.cmd("cq")
end

local failed = 0
local passed = 0

print("\n================== RUNNING NVIM-DOCX TEST SUITE ==================")
for _, name in ipairs(test_order) do
  local fn = tests[name]
  local ok, err = pcall(fn)
  if ok then
    passed = passed + 1
    print("  ✓ PASS: " .. name)
  else
    failed = failed + 1
    print("  ✗ FAIL: " .. name .. " -> " .. tostring(err))
  end
end

vim.fn.delete(test_dir, "rf")

print("==================================================================")
print(string.format("Result: %d passed, %d failed\n", passed, failed))

if failed > 0 then
  vim.cmd("cq")
else
  vim.cmd("qall!")
end
