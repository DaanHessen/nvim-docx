-- Live preview: renders the buffer as a document page in the browser and updates as you type.
-- A tiny HTTP server on 127.0.0.1 (libuv, no dependencies); the page polls for a new version.

local markdown = require("nvim-docx.markdown")
local config = require("nvim-docx.config")
local utils = require("nvim-docx.utils")

local M = {}

local uv = vim.uv or vim.loop

-- One server per buffer: { server, port, version, html, media_dirs, augroup }
local previews = {}

local function esc(s)
  return (tostring(s or ""):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"):gsub('"', "&quot;"))
end

local ALIGN_CSS = { left = "left", center = "center", centre = "center", right = "right", justify = "justify", both = "justify" }

local function css_color(value)
  if not value or value == false then
    return nil
  end
  local hex = require("nvim-docx.ooxml").color_hex(value)
  return hex and ("#" .. hex) or nil
end

local function run_style(fmt)
  local css = {}
  if fmt.font then
    css[#css + 1] = string.format("font-family:'%s'", esc(fmt.font))
  end
  local size = tonumber(fmt.size)
  if size then
    css[#css + 1] = string.format("font-size:%gpt", size)
  end
  local color = css_color(fmt.color)
  if color then
    css[#css + 1] = "color:" .. color
  end
  if fmt.highlight then
    css[#css + 1] = "background:" .. (css_color(fmt.highlight) or fmt.highlight)
  end
  return table.concat(css, ";")
end

local function media_url(src)
  if not src or src:match("^https?://") or src:match("^data:") then
    return src or ""
  end
  return "/media/" .. src:gsub("^%./", "")
end

local function inlines_html(inlines)
  local out = {}
  for _, it in ipairs(markdown.merge_inlines(inlines or {})) do
    if it.type == "text" then
      local s = esc(it.text):gsub("\n", "<br>")
      if it.code then
        s = "<code>" .. s .. "</code>"
      end
      if it.sup then
        s = "<sup>" .. s .. "</sup>"
      elseif it.sub then
        s = "<sub>" .. s .. "</sub>"
      end
      if it.strike then
        s = "<s>" .. s .. "</s>"
      end
      if it.underline then
        s = "<u>" .. s .. "</u>"
      end
      if it.italic then
        s = "<em>" .. s .. "</em>"
      end
      if it.bold then
        s = "<strong>" .. s .. "</strong>"
      end
      local style = run_style(it)
      if style ~= "" then
        s = string.format('<span style="%s">%s</span>', style, s)
      end
      out[#out + 1] = s
    elseif it.type == "link" then
      out[#out + 1] = string.format('<a href="%s">%s</a>', esc(it.url), inlines_html(it.inlines))
    elseif it.type == "image" then
      local w = it.width and string.format(' style="width:%s"', esc(it.width)) or ""
      out[#out + 1] = string.format('<img src="%s" alt="%s"%s>', esc(media_url(it.src)), esc(it.alt), w)
    elseif it.type == "footnote_ref" then
      out[#out + 1] = string.format('<sup class="fn">%s</sup>', esc(it.id))
    elseif it.type == "math" then
      out[#out + 1] = '<span class="math">' .. esc(it.text) .. "</span>"
    elseif it.type == "linebreak" then
      out[#out + 1] = "<br>"
    elseif it.type == "tab" then
      out[#out + 1] = '<span class="tab"></span>'
    end
  end
  return table.concat(out)
end

local function block_style(block)
  local kv = block.attrs and block.attrs.kv
  if not kv then
    return "", ""
  end
  local css = {}
  local align = kv.align and ALIGN_CSS[kv.align:lower()]
  if align then
    css[#css + 1] = "text-align:" .. align
  end
  local run = run_style({
    font = kv.font,
    size = block.type ~= "heading" and kv.size and kv.size:gsub("pt$", "") or nil,
    color = kv.color,
    highlight = kv.highlight,
  })
  if run ~= "" then
    css[#css + 1] = run
  end
  local class = kv.style and (" class=\"style-" .. esc(kv.style:lower():gsub("%s+", "-")) .. "\"") or ""
  return (#css > 0) and string.format(' style="%s"', table.concat(css, ";")) or "", class
end

function M.render_html(lines)
  local ast = markdown.markdown_to_ast(lines, config.options or {})
  local out = {}
  local list_stack = {} -- { tag, level }
  local function close_lists(to_level)
    while #list_stack > 0 and list_stack[#list_stack].level > to_level do
      out[#out + 1] = "</li></" .. table.remove(list_stack).tag .. ">"
    end
  end
  for _, b in ipairs(ast.blocks) do
    local style, class = block_style(b)
    if b.type ~= "list_item" then
      close_lists(-1)
    end
    if b.type == "heading" then
      local lvl = math.min(math.max(b.level or 1, 1), 6)
      out[#out + 1] = string.format("<h%d%s%s>%s</h%d>", lvl, style, class, inlines_html(b.inlines), lvl)
    elseif b.type == "paragraph" then
      out[#out + 1] = string.format("<p%s%s>%s</p>", style, class, inlines_html(b.inlines))
    elseif b.type == "blockquote" then
      out[#out + 1] = string.format("<blockquote%s><p>%s</p></blockquote>", style, inlines_html(b.inlines))
    elseif b.type == "list_item" then
      local lvl = b.level or 0
      local tag = b.ordered and "ol" or "ul"
      close_lists(lvl)
      local top = list_stack[#list_stack]
      if top and top.level == lvl and top.tag ~= tag then
        -- bullets followed by numbers (or the reverse) at the same level start a new list
        close_lists(lvl - 1)
        top = list_stack[#list_stack]
      end
      if not top or top.level < lvl then
        local start = (b.ordered and b.number and b.number ~= 1) and string.format(' start="%d"', b.number) or ""
        out[#out + 1] = string.format("<%s%s>", tag, start)
        list_stack[#list_stack + 1] = { tag = tag, level = lvl }
      else
        out[#out + 1] = "</li>"
      end
      local box = ""
      if b.checked == true then
        box = "&#9745; "
      elseif b.checked == false then
        box = "&#9744; "
      end
      out[#out + 1] = string.format("<li%s>%s%s", style, box, inlines_html(b.inlines))
    elseif b.type == "code_block" then
      out[#out + 1] = "<pre><code>" .. esc(b.text) .. "</code></pre>"
    elseif b.type == "math_block" then
      out[#out + 1] = '<div class="math block">' .. esc(b.text) .. "</div>"
    elseif b.type == "hr" then
      out[#out + 1] = "<hr>"
    elseif b.type == "pagebreak" then
      out[#out + 1] = '<div class="pagebreak"></div>'
    elseif b.type == "table" then
      local rows = { "<table" .. style .. ">" }
      local aligns = b.alignments or {}
      local function row(cells, tag)
        local r = { "<tr>" }
        for c, cell in ipairs(cells or {}) do
          r[#r + 1] = string.format('<%s style="text-align:%s">%s</%s>', tag, aligns[c] or "left", inlines_html(cell), tag)
        end
        r[#r + 1] = "</tr>"
        return table.concat(r)
      end
      rows[#rows + 1] = row(b.headers, "th")
      for _, r in ipairs(b.rows or {}) do
        rows[#rows + 1] = row(r, "td")
      end
      rows[#rows + 1] = "</table>"
      out[#out + 1] = table.concat(rows)
    end
  end
  close_lists(-1)
  if ast.footnotes and #ast.footnotes > 0 then
    out[#out + 1] = '<div class="footnotes"><hr>'
    for _, fn in ipairs(ast.footnotes) do
      out[#out + 1] = string.format("<p><sup>%s</sup> %s</p>", esc(fn.id), inlines_html(fn.inlines))
    end
    out[#out + 1] = "</div>"
  end
  return table.concat(out, "\n")
end

local PAGE = [[<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<style>
:root { color-scheme: light; }
body { margin: 0; background: #e8e8e8; font-family: Calibri, Carlito, "Segoe UI", Arial, sans-serif; }
#page { box-sizing: border-box; max-width: 8.5in; min-height: 11in; margin: 24px auto; padding: 1in;
  background: #fff; color: #000; box-shadow: 0 2px 12px rgba(0,0,0,.25); font-size: 11pt; line-height: 1.35; }
@media (max-width: 9in) { #page { margin: 0; padding: 16px; min-height: 100vh; box-shadow: none; } }
h1, h2, h3, h4, h5, h6 { color: #2F5496; margin: 1.1em 0 .4em; line-height: 1.2; }
h1 { font-size: 18pt; } h2 { font-size: 14pt; } h3 { font-size: 12pt; color: #1F3763; } h4, h5, h6 { font-size: 11pt; }
p { margin: 0 0 8pt; }
blockquote { margin: 0 0 8pt; padding-left: 12px; border-left: 3px solid #D0D7DE; color: #444; font-style: italic; }
.style-title { font-size: 28pt; margin-bottom: 4pt; }
.style-subtitle { font-size: 14pt; color: #595959; }
.style-quote { font-style: italic; color: #404040; text-align: center; margin: 10pt .6in; }
pre { background: #f5f5f5; padding: 8px; overflow-x: auto; font-size: 10pt; }
code { font-family: Consolas, "Courier New", monospace; }
table { border-collapse: collapse; margin: 0 0 10pt; }
th, td { border: 1px solid #D0D7DE; padding: 3px 8px; vertical-align: top; }
th { font-weight: normal; }
img { max-width: 100%%; height: auto; }
hr { border: 0; border-top: 1px solid #999; }
.pagebreak { border-top: 2px dashed #bbb; margin: 18px -1in; }
.tab { display: inline-block; width: .5in; }
.math { font-family: "Cambria Math", "STIX Two Math", serif; font-style: italic; }
.math.block { text-align: center; margin: 8pt 0; }
.footnotes { font-size: 9pt; margin-top: 2em; }
#status { position: fixed; right: 10px; bottom: 8px; font: 11px sans-serif; color: #888; }
</style></head>
<body><div id="page">%s</div><div id="status">nvim-docx live preview</div>
<script>
let version = %d;
async function poll() {
  try {
    const r = await fetch("/version", { cache: "no-store" });
    const v = parseInt(await r.text(), 10);
    if (v !== version) {
      const h = await fetch("/body", { cache: "no-store" });
      const y = window.scrollY;
      document.getElementById("page").innerHTML = await h.text();
      window.scrollTo(0, y);
      version = v;
    }
    document.getElementById("status").textContent = "nvim-docx live preview";
  } catch (e) {
    document.getElementById("status").textContent = "preview stopped (Neovim closed the server)";
  }
  setTimeout(poll, 400);
}
poll();
</script></body></html>]]

local MIME = {
  png = "image/png", jpg = "image/jpeg", jpeg = "image/jpeg", gif = "image/gif", bmp = "image/bmp",
  svg = "image/svg+xml", webp = "image/webp", emf = "image/x-emf", wmf = "image/x-wmf", tif = "image/tiff",
  tiff = "image/tiff",
}

local function respond(client, status, ctype, body)
  local head = string.format(
    "HTTP/1.1 %s\r\nContent-Type: %s\r\nContent-Length: %d\r\nCache-Control: no-store\r\nConnection: close\r\n\r\n",
    status, ctype, #body)
  client:write(head .. body, function()
    client:close()
  end)
end

local function read_file(path)
  local f = io.open(path, "rb")
  if not f then
    return nil
  end
  local data = f:read("*a")
  f:close()
  return data
end

local function handle(p, client, request)
  local path = request:match("^GET (%S+)") or "/"
  path = path:gsub("%?.*$", "")
  if path == "/" then
    respond(client, "200 OK", "text/html; charset=utf-8", string.format(PAGE, esc(p.title), p.html, p.version))
  elseif path == "/favicon.ico" then
    respond(client, "204 No Content", "text/plain", "")
  elseif path == "/version" then
    respond(client, "200 OK", "text/plain", tostring(p.version))
  elseif path == "/body" then
    respond(client, "200 OK", "text/html; charset=utf-8", p.html)
  elseif path:match("^/media/") then
    local rel = path:sub(8):gsub("%%(%x%x)", function(h)
      return string.char(tonumber(h, 16))
    end)
    if rel:find("..", 1, true) then
      return respond(client, "403 Forbidden", "text/plain", "forbidden")
    end
    for _, dir in ipairs(p.media_dirs) do
      local data = read_file(dir .. "/" .. rel)
      if data then
        local ext = (rel:match("%.(%w+)$") or ""):lower()
        return respond(client, "200 OK", MIME[ext] or "application/octet-stream", data)
      end
    end
    respond(client, "404 Not Found", "text/plain", "not found")
  else
    respond(client, "404 Not Found", "text/plain", "not found")
  end
end

local function open_browser(url)
  if vim.ui and vim.ui.open then
    local ok = pcall(vim.ui.open, url)
    if ok then
      return
    end
  end
  local cmd
  if vim.fn.has("mac") == 1 then
    cmd = { "open", url }
  elseif vim.fn.has("win32") == 1 then
    cmd = { "cmd.exe", "/c", "start", "", url }
  else
    cmd = { "xdg-open", url }
  end
  vim.fn.jobstart(cmd, { detach = true })
end

local function refresh(bufnr)
  local p = previews[bufnr]
  if not p or not vim.api.nvim_buf_is_valid(bufnr) then
    return
  end
  local ok, html = pcall(M.render_html, vim.api.nvim_buf_get_lines(bufnr, 0, -1, false))
  p.html = ok and html or ("<pre>" .. esc(html) .. "</pre>")
  p.version = p.version + 1
end

function M.stop(bufnr)
  bufnr = bufnr or vim.api.nvim_get_current_buf()
  local p = previews[bufnr]
  if not p then
    return false
  end
  previews[bufnr] = nil
  pcall(vim.api.nvim_del_augroup_by_id, p.augroup)
  if p.timer then
    p.timer:stop()
    p.timer:close()
  end
  if not p.server:is_closing() then
    p.server:close()
  end
  return true
end

function M.start(bufnr)
  bufnr = bufnr or vim.api.nvim_get_current_buf()
  if previews[bufnr] then
    open_browser(previews[bufnr].url)
    return previews[bufnr].url
  end

  local name = vim.api.nvim_buf_get_name(bufnr)
  local state = vim.b[bufnr].nvim_docx_state
  local media_dirs = {}
  if state and state.workspace then
    -- media from the document is only extracted on demand
    local init = require("nvim-docx")
    if init.extract_media then
      init.extract_media(bufnr)
    end
    media_dirs[#media_dirs + 1] = state.workspace
  end
  if name ~= "" then
    media_dirs[#media_dirs + 1] = vim.fn.fnamemodify(name, ":p:h")
  end

  local server = uv.new_tcp()
  local ok, err = pcall(function()
    assert(server:bind("127.0.0.1", 0))
  end)
  if not ok then
    utils.notify("Preview server failed to start: " .. tostring(err), "error")
    return nil
  end
  local port = server:getsockname().port
  local p = {
    server = server,
    port = port,
    url = string.format("http://127.0.0.1:%d/", port),
    version = 0,
    html = "",
    media_dirs = media_dirs,
    title = vim.fn.fnamemodify(name, ":t"),
  }
  previews[bufnr] = p

  server:listen(64, function(listen_err)
    if listen_err then
      return
    end
    local client = uv.new_tcp()
    server:accept(client)
    local chunks = {}
    client:read_start(function(read_err, data)
      if read_err or not data then
        client:close()
        return
      end
      chunks[#chunks + 1] = data
      local request = table.concat(chunks)
      if request:find("\r\n\r\n", 1, true) then
        client:read_stop()
        handle(p, client, request)
      end
    end)
  end)

  refresh(bufnr)
  p.timer = uv.new_timer()
  p.augroup = vim.api.nvim_create_augroup("NvimDocxPreview" .. bufnr, { clear = true })
  vim.api.nvim_create_autocmd({ "TextChanged", "TextChangedI", "BufWritePost" }, {
    group = p.augroup,
    buffer = bufnr,
    callback = function()
      -- debounce: re-render once typing pauses
      p.timer:stop()
      p.timer:start(250, 0, vim.schedule_wrap(function()
        refresh(bufnr)
      end))
    end,
  })
  vim.api.nvim_create_autocmd({ "BufWipeout", "BufDelete" }, {
    group = p.augroup,
    buffer = bufnr,
    callback = function()
      M.stop(bufnr)
    end,
  })

  open_browser(p.url)
  utils.notify("Live preview at " .. p.url .. " (:DocxPreviewStop to end)", "info")
  return p.url
end

function M.stop_all()
  for bufnr in pairs(previews) do
    M.stop(bufnr)
  end
end

return M
