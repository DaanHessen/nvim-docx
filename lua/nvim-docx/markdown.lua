-- Pure-Lua Markdown parser and renderer for nvim-docx AST.

local M = {}

-- Run formatting carried by text nodes. font/size/color/highlight come from attribute spans
-- ([text]{font="Georgia" size=14 color=#C00000 highlight=yellow}); size is in points.
local FMT_FLAGS = { "bold", "italic", "strike", "underline", "code", "sup", "sub" }
local FMT_ATTRS = { "font", "size", "color", "highlight" }
M.FMT_ATTRS = FMT_ATTRS

local function clone_fmt(fmt)
  local out = {}
  fmt = fmt or {}
  for _, k in ipairs(FMT_FLAGS) do
    out[k] = fmt[k] or false
  end
  for _, k in ipairs(FMT_ATTRS) do
    out[k] = fmt[k]
  end
  return out
end

local function same_fmt(a, b)
  for _, k in ipairs(FMT_FLAGS) do
    if (a[k] or false) ~= (b[k] or false) then
      return false
    end
  end
  for _, k in ipairs(FMT_ATTRS) do
    if a[k] ~= b[k] then
      return false
    end
  end
  return true
end

-- Parse a pandoc attribute block body ("font=\"Georgia\" size=14 .underline") into classes and keys
function M.parse_attrs(body)
  local attrs = { classes = {}, kv = {} }
  local pos = 1
  local len = #body
  while pos <= len do
    local ws = body:match("^%s+", pos)
    if ws then
      pos = pos + #ws
    end
    if pos > len then
      break
    end
    local cls = body:match("^%.([%w%-_]+)", pos)
    if cls then
      attrs.classes[cls] = true
      pos = pos + 1 + #cls
    else
      local key = body:match("^([%w%-_]+)=", pos)
      if not key then
        return nil
      end
      pos = pos + #key + 1
      local q = body:sub(pos, pos)
      local val
      if q == '"' or q == "'" then
        local close = body:find(q, pos + 1, true)
        if not close then
          return nil
        end
        val = body:sub(pos + 1, close - 1)
        pos = close + 1
      else
        val = body:match("^[^%s}]+", pos) or ""
        pos = pos + #val
      end
      attrs.kv[key] = val
    end
  end
  return attrs
end

-- Run formatting from parsed attributes; nil when the attributes say nothing about runs
local function fmt_from_attrs(attrs, base)
  if not attrs then
    return nil
  end
  local f = clone_fmt(base)
  local any = false
  local kv, cls = attrs.kv, attrs.classes
  if kv.font or kv["font-family"] then
    f.font = kv.font or kv["font-family"]
    any = true
  end
  local size = kv.size or kv["font-size"]
  if size then
    local n = tonumber((size:gsub("pt$", "")))
    if n and n > 0 then
      f.size = n
      any = true
    end
  end
  if kv.color then
    f.color = kv.color
    any = true
  end
  if kv.highlight or cls.mark then
    f.highlight = kv.highlight or "yellow"
    any = true
  end
  if cls.underline or cls.ul then
    f.underline = true
    any = true
  end
  if cls.smallcaps then
    any = true
  end
  return any and f or nil
end
M.fmt_from_attrs = fmt_from_attrs

local function format_number(n)
  if n == math.floor(n) then
    return tostring(math.floor(n))
  end
  return tostring(n)
end

-- Canonical attribute text for a run's font/size/color/highlight (nil when it has none)
function M.span_attrs(item)
  local parts = {}
  if item.font then
    parts[#parts + 1] = string.format('font="%s"', item.font)
  end
  if item.size then
    parts[#parts + 1] = "size=" .. format_number(item.size)
  end
  if item.color then
    parts[#parts + 1] = "color=" .. item.color
  end
  if item.highlight then
    parts[#parts + 1] = "highlight=" .. item.highlight
  end
  if #parts == 0 then
    return nil
  end
  return "{" .. table.concat(parts, " ") .. "}"
end

function M.merge_inlines(inlines)
  local merged = {}
  for _, item in ipairs(inlines or {}) do
    if item.type == "text" then
      if item.text and item.text ~= "" then
        local prev = merged[#merged]
        if prev and prev.type == "text" and same_fmt(prev, item) then
          prev.text = prev.text .. item.text
        else
          local copy = clone_fmt(item)
          copy.type = "text"
          copy.text = item.text
          merged[#merged + 1] = copy
        end
      end
    else
      merged[#merged + 1] = item
    end
  end
  return merged
end

--------------------------------------------------------------------------------
-- Inline Parser (Markdown string -> AST inlines)
--------------------------------------------------------------------------------

-- Helper to match balanced [...] starting at `pos` (where str:sub(pos, pos) == "[")
-- Drop leading/trailing tabs, line breaks and spaces of a paragraph (Word keeps them, Markdown cannot)
-- and the whitespace around interior line breaks. Leading spaces are kept unless trim_leading_spaces.
function M.trim_inlines(inlines, trim_leading_spaces)
  if not inlines then
    return {}
  end
  while #inlines > 0 do
    local first = inlines[1]
    if first.type == "tab" or first.type == "linebreak" then
      table.remove(inlines, 1)
    elseif trim_leading_spaces and first.type == "text" then
      first.text = (first.text or ""):gsub("^[ \t]+", "")
      if first.text == "" then
        table.remove(inlines, 1)
      else
        break
      end
    else
      break
    end
  end
  while #inlines > 0 do
    local last = inlines[#inlines]
    if last.type == "tab" or last.type == "linebreak" then
      table.remove(inlines, #inlines)
    elseif last.type == "text" then
      last.text = (last.text or ""):gsub("[ \t]+$", "")
      if last.text == "" then
        table.remove(inlines, #inlines)
      else
        break
      end
    else
      break
    end
  end
  -- Clean up spaces/tabs adjacent to interior linebreaks and collapse consecutive linebreaks
  local i = 1
  while i <= #inlines do
    if inlines[i].type == "linebreak" then
      -- Trim trailing spaces/tabs before linebreak
      if i > 1 then
        local prev = inlines[i - 1]
        if prev.type == "tab" then
          table.remove(inlines, i - 1)
          i = i - 1
        elseif prev.type == "text" then
          prev.text = (prev.text or ""):gsub("[ \t]+$", "")
          if prev.text == "" then
            table.remove(inlines, i - 1)
            i = i - 1
          end
        end
      end
      -- Trim leading spaces/tabs or duplicate linebreaks after linebreak
      while i + 1 <= #inlines do
        local nxt = inlines[i + 1]
        if nxt.type == "linebreak" or nxt.type == "tab" then
          table.remove(inlines, i + 1)
        elseif nxt.type == "text" then
          nxt.text = (nxt.text or ""):gsub("^[ \t]+", "")
          if nxt.text == "" then
            table.remove(inlines, i + 1)
          else
            break
          end
        else
          break
        end
      end
    end
    i = i + 1
  end
  -- Re-check leading/trailing linebreak in case interior cleanup exposed one
  while #inlines > 0 and (inlines[1].type == "linebreak" or inlines[1].type == "tab") do
    table.remove(inlines, 1)
  end
  while #inlines > 0 and (inlines[#inlines].type == "linebreak" or inlines[#inlines].type == "tab") do
    table.remove(inlines, #inlines)
  end
  return M.merge_inlines(inlines)
end

local function match_bracketed(str, pos)
  if str:sub(pos, pos) ~= "[" then
    return nil, nil
  end
  local depth = 0
  local i = pos
  local len = #str
  while i <= len do
    local ch = str:sub(i, i)
    if ch == "\\" then
      i = i + 2
    elseif ch == "[" then
      depth = depth + 1
      i = i + 1
    elseif ch == "]" then
      depth = depth - 1
      if depth == 0 then
        return str:sub(pos + 1, i - 1), i
      end
      i = i + 1
    else
      i = i + 1
    end
  end
  return nil, nil
end

-- Helper to match link/image destination (url "title") starting at `pos` (where str:sub(pos, pos) == "(")
local function match_parens_dest(str, pos)
  if str:sub(pos, pos) ~= "(" then
    return nil, nil, nil
  end
  local depth = 0
  local i = pos
  local len = #str
  local in_quote = nil
  while i <= len do
    local ch = str:sub(i, i)
    if ch == "\\" then
      i = i + 2
    elseif in_quote then
      if ch == in_quote then
        in_quote = nil
      end
      i = i + 1
    else
      if ch == '"' or ch == "'" then
        in_quote = ch
        i = i + 1
      elseif ch == "(" then
        depth = depth + 1
        i = i + 1
      elseif ch == ")" then
        depth = depth - 1
        if depth == 0 then
          local inner = str:sub(pos + 1, i - 1):gsub("^%s+", ""):gsub("%s+$", "")
          local url, title = inner:match("^(%S+)%s+[\"'](.-)[\"']$")
          if not url then
            url = inner
          end
          if url:sub(1, 1) == "<" and url:sub(-1) == ">" then
            url = url:sub(2, -2)
          end
          return url, title, i
        end
        i = i + 1
      else
        i = i + 1
      end
    end
  end
  return nil, nil, nil
end

-- Parse optional Pandoc attribute block `{width="..." height="..."}` starting at `pos`
local function match_image_attrs(str, pos)
  if str:sub(pos, pos) ~= "{" then
    return nil, nil, pos - 1
  end
  local close = str:find("}", pos + 1, true)
  if not close then
    return nil, nil, pos - 1
  end
  local body = str:sub(pos + 1, close - 1)
  local w = body:match('width%s*=%s*"([^"]+)"') or body:match("width%s*=%s*'([^']+)'") or body:match("width%s*=%s*([^%s}]+)")
  local h = body:match('height%s*=%s*"([^"]+)"') or body:match("height%s*=%s*'([^']+)'") or body:match("height%s*=%s*([^%s}]+)")
  return w, h, close
end

local function is_alnum(ch)
  return ch ~= "" and ch:match("^[%w]$") ~= nil
end

-- ^sup^ / ~sub~ starting at pos: content may not contain unescaped whitespace (pandoc rules; a space
-- is written as "\ "). Returns start, end, content.
local function find_script(str, pos, delim)
  local i = pos + 1
  local len = #str
  while i <= len do
    local c = str:sub(i, i)
    if c == "\\" then
      i = i + 2
    elseif c:match("%s") then
      return nil
    elseif c == delim then
      if i == pos + 1 then
        return nil
      end
      return pos, i, str:sub(pos + 1, i - 1)
    else
      i = i + 1
    end
  end
  return nil
end

function M.parse_inlines(str, fmt)
  fmt = clone_fmt(fmt)
  local out = {}
  if not str or str == "" then
    return out
  end

  local len = #str
  local pos = 1
  local buf = {}

  local function flush_buf()
    if #buf > 0 then
      local node = clone_fmt(fmt)
      node.type = "text"
      node.text = table.concat(buf)
      out[#out + 1] = node
      buf = {}
    end
  end

  while pos <= len do
    local ch = str:sub(pos, pos)

    -- 1. Escaped punctuation
    if ch == "\\" and pos < len then
      local next_ch = str:sub(pos + 1, pos + 1)
      if next_ch:match("^[\\`%*_%{%}%[%]%(%)#%+%-%.,!|~^<>$:= ]$") then
        buf[#buf + 1] = next_ch
        pos = pos + 2
      else
        buf[#buf + 1] = ch
        pos = pos + 1
      end

    -- 2. HTML <br> / <br/>, <u>...</u>, or autolinks <https://...>
    elseif ch == "<" then
      local br_s, br_e = str:find("^<[bB][rR]%s*/?>", pos)
      if br_s then
        flush_buf()
        out[#out + 1] = { type = "linebreak" }
        pos = br_e + 1
      else
        local u_s, u_e, u_inner = str:find("^<[uU]>(.-)</[uU]>", pos)
        if u_s and u_inner ~= "" then
          flush_buf()
          local sub_fmt = clone_fmt(fmt)
          sub_fmt.underline = true
          for _, n in ipairs(M.parse_inlines(u_inner, sub_fmt)) do
            out[#out + 1] = n
          end
          pos = u_e + 1
        else
          local auto_s, auto_e, auto_url = str:find("^<(https?://[^>%s]+)>", pos)
          if not auto_s then
            auto_s, auto_e, auto_url = str:find("^<(mailto:[^>%s]+)>", pos)
          end
          if auto_s then
            flush_buf()
            local link_text = clone_fmt(fmt)
            link_text.type = "text"
            link_text.text = auto_url:gsub("^mailto:", "")
            out[#out + 1] = {
              type = "link",
              url = auto_url,
              inlines = { link_text },
            }
            pos = auto_e + 1
          else
            buf[#buf + 1] = ch
            pos = pos + 1
          end
        end
      end

    -- 3. Inline code span `...`
    elseif ch == "`" then
      local _, tick_end, ticks = str:find("^(`+)", pos)
      local close_s, close_e = str:find(ticks, tick_end + 1, true)
      if close_s then
        flush_buf()
        local code_text = str:sub(tick_end + 1, close_s - 1)
        if #code_text >= 2 and code_text:sub(1, 1) == " " and code_text:sub(-1) == " " and code_text:find("%S") then
          code_text = code_text:sub(2, -2)
        end
        local node = clone_fmt(fmt)
        node.type = "text"
        node.code = true
        node.text = code_text
        out[#out + 1] = node
        pos = close_e + 1
      else
        buf[#buf + 1] = ticks
        pos = tick_end + 1
      end

    -- 4. Image ![alt](src){attrs}
    elseif ch == "!" and str:sub(pos + 1, pos + 1) == "[" then
      local alt, bracket_end = match_bracketed(str, pos + 1)
      if alt and bracket_end and str:sub(bracket_end + 1, bracket_end + 1) == "(" then
        local src, title, paren_end = match_parens_dest(str, bracket_end + 1)
        if src and paren_end then
          flush_buf()
          local w, h, attr_end = match_image_attrs(str, paren_end + 1)
          out[#out + 1] = {
            type = "image",
            alt = alt,
            src = src,
            title = title,
            width = w,
            height = h,
          }
          pos = attr_end + 1
        else
          buf[#buf + 1] = ch
          pos = pos + 1
        end
      else
        buf[#buf + 1] = ch
        pos = pos + 1
      end

    -- 5. Footnote ref [^id], Link [label](url), or Underlined span [label]{.underline}
    elseif ch == "[" then
      local fn_s, fn_e, fn_id = str:find("^%[%^([^%]%s]+)%]", pos)
      if fn_s and str:sub(fn_e + 1, fn_e + 1) ~= ":" then
        flush_buf()
        out[#out + 1] = {
          type = "footnote_ref",
          id = fn_id,
        }
        pos = fn_e + 1
      else
        local label, bracket_end = match_bracketed(str, pos)
        if label and bracket_end and str:sub(bracket_end + 1, bracket_end + 1) == "(" then
          local url, title, paren_end = match_parens_dest(str, bracket_end + 1)
          if url and paren_end then
            flush_buf()
            local child_inlines = M.parse_inlines(label, fmt)
            if #child_inlines == 0 then
              local tnode = clone_fmt(fmt)
              tnode.type = "text"
              tnode.text = url
              child_inlines = { tnode }
            end
            out[#out + 1] = {
              type = "link",
              url = url,
              title = title,
              inlines = child_inlines,
            }
            pos = paren_end + 1
          else
            buf[#buf + 1] = ch
            pos = pos + 1
          end
        elseif label and bracket_end then
          local a_s, a_e, a_body = str:find("^%{([^}]*)%}", bracket_end + 1)
          local sub_fmt = a_s and label ~= "" and fmt_from_attrs(M.parse_attrs(a_body), fmt) or nil
          if sub_fmt then
            flush_buf()
            for _, n in ipairs(M.parse_inlines(label, sub_fmt)) do
              out[#out + 1] = n
            end
            pos = a_e + 1
          else
            buf[#buf + 1] = ch
            pos = pos + 1
          end
        else
          buf[#buf + 1] = ch
          pos = pos + 1
        end
      end

    -- 6. Bold/Italic with * or _
    elseif ch == "*" or ch == "_" then
      local prev_ch = pos > 1 and str:sub(pos - 1, pos - 1) or ""
      local pat = (ch == "*") and "^(%*+)" or "^(_+)"
      local _, run_end, delim = str:find(pat, pos)
      local dlen = #delim
      local after_open = str:sub(run_end + 1, run_end + 1)
      local can_open = (after_open ~= "" and not after_open:find("^%s$"))
      if ch == "_" and is_alnum(prev_ch) then
        can_open = false
      end

      if can_open and (dlen == 3 or dlen == 2 or dlen == 1) then
        local close_s, close_e = nil, nil
        local search_pos = run_end + 1
        local run_pat = (ch == "*") and "(%*+)" or "(_+)"
        while search_pos <= len do
          local rs, re = str:find(run_pat, search_pos)
          if not rs then break end
          local rlen = re - rs + 1
          local before = str:sub(rs - 1, rs - 1)
          local after = str:sub(re + 1, re + 1)
          local valid_close = (before ~= "\\" and not before:find("^%s$"))
          if ch == "_" and is_alnum(after) then
            valid_close = false
          end

          if valid_close then
            if rlen == dlen then
              close_s = rs
              close_e = re
              break
            elseif rlen >= 3 and rlen > dlen then
              -- Consume the trailing `dlen` characters of this run as the outer closing delimiter,
              -- leaving the leading `rlen - dlen` characters inside `inner` to close inner spans!
              close_s = re - dlen + 1
              close_e = re
              break
            end
          end
          search_pos = re + 1
        end

        if close_s and close_s > run_end + 1 then
          flush_buf()
          local sub_fmt = clone_fmt(fmt)
          if dlen == 3 then
            sub_fmt.bold = true
            sub_fmt.italic = true
          elseif dlen == 2 then
            sub_fmt.bold = true
          else
            sub_fmt.italic = true
          end
          local inner = str:sub(run_end + 1, close_s - 1)
          local inner_nodes = M.parse_inlines(inner, sub_fmt)
          for _, n in ipairs(inner_nodes) do
            out[#out + 1] = n
          end
          pos = close_e + 1
        else
          buf[#buf + 1] = delim
          pos = run_end + 1
        end
      else
        buf[#buf + 1] = delim
        pos = run_end + 1
      end

    -- 7. Strikethrough ~~...~~ or Subscript ~...~
    elseif ch == "~" then
      if str:sub(pos, pos + 1) == "~~" then
        local close_s, close_e = str:find("~~", pos + 2, true)
        if close_s and close_s > pos + 2 then
          flush_buf()
          local sub_fmt = clone_fmt(fmt)
          sub_fmt.strike = true
          local inner = str:sub(pos + 2, close_s - 1)
          for _, n in ipairs(M.parse_inlines(inner, sub_fmt)) do
            out[#out + 1] = n
          end
          pos = close_e + 1
        else
          buf[#buf + 1] = "~~"
          pos = pos + 2
        end
      else
        local sub_s, sub_e, sub_txt = find_script(str, pos, "~")
        if sub_s then
          flush_buf()
          local sub_fmt = clone_fmt(fmt)
          sub_fmt.sub = true
          for _, n in ipairs(M.parse_inlines(sub_txt, sub_fmt)) do
            out[#out + 1] = n
          end
          pos = sub_e + 1
        else
          buf[#buf + 1] = ch
          pos = pos + 1
        end
      end

    -- 8. Superscript ^...^
    elseif ch == "^" then
      local sup_s, sup_e, sup_txt = find_script(str, pos, "^")
      if sup_s then
        flush_buf()
        local sub_fmt = clone_fmt(fmt)
        sub_fmt.sup = true
        for _, n in ipairs(M.parse_inlines(sup_txt, sub_fmt)) do
          out[#out + 1] = n
        end
        pos = sup_e + 1
      else
        buf[#buf + 1] = ch
        pos = pos + 1
      end

    -- 9. Inline Math $...$
    elseif ch == "$" and str:sub(pos + 1, pos + 1) ~= "$" then
      local after_open = str:sub(pos + 1, pos + 1)
      if after_open ~= "" and not after_open:find("^%s$") then
        local close_s = nil
        local sp = pos + 1
        while sp <= len do
          local s = str:find("$", sp, true)
          if not s then break end
          local prev = str:sub(s - 1, s - 1)
          local nxt = str:sub(s + 1, s + 1)
          if prev ~= "\\" and not prev:find("^%s$") and not nxt:match("^%d$") then
            close_s = s
            break
          end
          sp = s + 1
        end
        if close_s and close_s > pos + 1 then
          flush_buf()
          out[#out + 1] = {
            type = "math",
            text = str:sub(pos + 1, close_s - 1),
          }
          pos = close_s + 1
        else
          buf[#buf + 1] = ch
          pos = pos + 1
        end
      else
        buf[#buf + 1] = ch
        pos = pos + 1
      end

    -- 10. Tab character
    elseif ch == "\t" then
      flush_buf()
      out[#out + 1] = { type = "tab" }
      pos = pos + 1

    else
      buf[#buf + 1] = ch
      pos = pos + 1
    end
  end

  flush_buf()
  return M.merge_inlines(out)
end

--------------------------------------------------------------------------------
-- Inline Renderer (AST inlines -> Markdown string)
--------------------------------------------------------------------------------

-- Backslash-escape characters in plain document text that the inline parser would otherwise read as
-- Markdown syntax, so text such as "*", "$5 and $10" or "a_b_c" survives a round trip unchanged
local function escape_inline_text(txt)
  local out = {}
  local len = #txt
  for i = 1, len do
    local ch = txt:sub(i, i)
    local nxt = txt:sub(i + 1, i + 1)
    if ch == "\\" or ch == "*" or ch == "`" or ch == "$" or ch == "~" or ch == "^" then
      out[#out + 1] = "\\" .. ch
    elseif ch == "_" then
      -- intraword underscores (snake_case) are never emphasis
      local prev = txt:sub(i - 1, i - 1)
      if prev:match("%w") and nxt:match("%w") then
        out[#out + 1] = ch
      else
        out[#out + 1] = "\\_"
      end
    elseif ch == "[" then
      -- only brackets that would form a link, image, footnote or attribute span need escaping
      local close = txt:find("]", i + 1, true)
      local after = close and txt:sub(close + 1, close + 1) or ""
      if nxt == "^" or after == "(" or after == "{" or after == "[" then
        out[#out + 1] = "\\["
      else
        out[#out + 1] = ch
      end
    elseif ch == "<" and nxt:match("[%a/!]") then
      out[#out + 1] = "\\<"
    else
      out[#out + 1] = ch
    end
  end
  return table.concat(out)
end

M.escape_inline_text = escape_inline_text

-- Escape a rendered paragraph whose first characters would start a different block (list, heading,
-- quote, table, fence, footnote definition, horizontal rule)
local function escape_block_start(text)
  local lead, rest = text:match("^(%s*)(.*)$")
  if rest == "" then
    return text
  end
  local num, delim = rest:match("^(%d+)([.)])%s")
  if num then
    return lead .. num .. "\\" .. rest:sub(#num + 1)
  end
  if rest:match("^[#>|+-]") or rest:match("^[=-][=-]+%s*$") or rest:match("^:::") then
    return lead .. "\\" .. rest
  end
  return text
end

M.escape_block_start = escape_block_start

local function code_span(txt)
  if txt:find("`", 1, true) then
    local pad = (txt:sub(1, 1) == "`" or txt:sub(-1) == "`") and " " or ""
    return "``" .. pad .. txt .. pad .. "``"
  end
  return "`" .. txt .. "`"
end

-- Emphasis markers are opened and closed as a stack across consecutive text runs, so adjacent runs
-- with overlapping formatting render as nested Markdown ("**(*Rapport*)**") that parses back to the
-- same runs, instead of colliding markers ("**(*****Rapport**").
local MARK_ORDER = { "bold", "italic", "strike", "underline", "sup", "sub" }

function M.render_inlines(inlines, opts)
  opts = opts or {}
  local merged = M.merge_inlines(inlines or {})
  local parts = {}
  local in_table = opts.in_table or false
  local format = opts.markdown_format or "markdown"

  local stack = {}
  local pending_ws = ""
  local opening_now = false
  local closes_with_bold = false
  local function out(str)
    if str ~= "" then
      parts[#parts + 1] = str
    end
  end
  local function close_from(idx)
    for k = #stack, idx, -1 do
      out(stack[k].close)
      stack[k] = nil
    end
  end
  local function flush_ws()
    out(pending_ws)
    pending_ws = ""
  end
  local function close_all()
    close_from(1)
    flush_ws()
  end
  local function marker(key, core)
    -- next to a literal (escaped) asterisk, * markers would be misread; use _ instead
    local starry = core:sub(1, 2) == "\\*" or core:sub(-1) == "*"
    if key == "bold" then
      if starry then
        return "__", "__"
      end
      return "**", "**"
    elseif key == "italic" then
      -- "*" right after a just-closed "**" would read as "***"; one opened together with bold is fine
      local prev = parts[#parts]
      if starry or (prev and prev:sub(-1) == "*" and not (opening_now and closes_with_bold)) then
        return "_", "_"
      end
      return "*", "*"
    elseif key == "strike" then
      return "~~", "~~"
    elseif key == "underline" then
      return "<u>", "</u>"
    elseif key == "sup" then
      return "^", "^"
    end
    return "~", "~"
  end
  -- Does italic starting at merged[idx] end at the same run as the bold around it? ("***x***" is fine,
  -- "***x*.**" is not)
  local function italic_closes_with_bold(idx)
    for j = idx, #merged do
      local it = merged[j]
      if it.type ~= "text" then
        return true
      end
      local b, i = it.bold or false, it.italic or false
      if b ~= i then
        return false
      end
      if not b then
        return true
      end
    end
    return true
  end

  local function emit_text(item, idx)
    closes_with_bold = italic_closes_with_bold(idx)
    local txt = item.text or ""
    local lead, core, trail
    if item.code then
      lead, core, trail = "", code_span(txt), ""
    else
      lead, core, trail = escape_inline_text(txt):match("^(%s*)(.-)(%s*)$")
    end
    if in_table then
      core = core:gsub("|", "\\|"):gsub("\r?\n", "<br>")
      lead = lead:gsub("\r?\n", "<br>")
      trail = trail:gsub("\r?\n", "<br>")
    end
    if (item.sup or item.sub) and not item.code then
      core = core:gsub("(%s)", "\\%1")
    end
    if core == "" then
      -- whitespace between runs keeps the surrounding formatting open
      pending_ws = pending_ws .. lead .. trail
      return
    end
    local span = M.span_attrs(item)
    for idx, m in ipairs(stack) do
      if (m.key == "span" and m.attrs ~= span) or (m.key ~= "span" and not item[m.key]) then
        close_from(idx)
        break
      end
    end
    flush_ws()
    out(lead)
    local open = {}
    for _, m in ipairs(stack) do
      open[m.key] = true
    end
    if span and not open.span then
      -- an attribute span must enclose the emphasis inside it, so reopen everything within it
      close_from(1)
      open = {}
      out("[")
      stack[#stack + 1] = { key = "span", attrs = span, close = "]" .. span }
    end
    opening_now = false
    for _, key in ipairs(MARK_ORDER) do
      if item[key] and not open[key] and not (key == "sub" and item.sup) then
        local o, c = marker(key, core)
        out(o)
        stack[#stack + 1] = { key = key, close = c }
        opening_now = true
      end
    end
    out(core)
    pending_ws = trail
  end

  for item_idx, item in ipairs(merged) do
    if item.type ~= "text" and item.type ~= "tab" and item.type ~= "footnote_ref" then
      close_all()
    elseif item.type == "footnote_ref" then
      flush_ws()
    end
    if item.type == "text" then
      emit_text(item, item_idx)
    elseif item.type == "link" then
      -- link text is blue and underlined by the Hyperlink style; that is not formatting of its own
      local label_inlines = {}
      for k, il in ipairs(item.inlines or {}) do
        label_inlines[k] = (il.underline or il.color) and vim.tbl_extend("force", il, { underline = false, color = false }) or il
      end
      local label = M.render_inlines(label_inlines, opts)
      if label == "" then
        label = item.url or ""
      end
      if item.url and item.url ~= "" then
        if item.title and item.title ~= "" then
          parts[#parts + 1] = string.format('[%s](%s "%s")', label, item.url, item.title)
        else
          parts[#parts + 1] = string.format("[%s](%s)", label, item.url)
        end
      else
        parts[#parts + 1] = label
      end
    elseif item.type == "image" then
      local alt = item.alt or ""
      local src = item.src or ""
      local base
      if item.title and item.title ~= "" then
        base = string.format('![%s](%s "%s")', alt, src, item.title)
      else
        base = string.format("![%s](%s)", alt, src)
      end
      if format ~= "gfm" and (item.width or item.height) then
        local attrs = {}
        if item.width then attrs[#attrs + 1] = string.format('width="%s"', item.width) end
        if item.height then attrs[#attrs + 1] = string.format('height="%s"', item.height) end
        base = base .. "{" .. table.concat(attrs, " ") .. "}"
      end
      parts[#parts + 1] = base
    elseif item.type == "footnote_ref" then
      parts[#parts + 1] = string.format("[^%s]", tostring(item.id or "1"))
    elseif item.type == "math" then
      parts[#parts + 1] = "$" .. (item.text or "") .. "$"
    elseif item.type == "linebreak" then
      if in_table or opts.single_line then
        parts[#parts + 1] = "<br>"
      else
        parts[#parts + 1] = "\\\n"
      end
    elseif item.type == "tab" then
      -- a tab is whitespace: markers must not close right after it
      pending_ws = pending_ws .. "\t"
    end
  end

  close_all()
  return table.concat(parts)
end

--------------------------------------------------------------------------------
-- Block Parser (Markdown lines -> Document AST)
--------------------------------------------------------------------------------

local function split_pipe_row(line)
  local trimmed = line:gsub("^%s+", ""):gsub("%s+$", "")
  if trimmed:sub(1, 1) == "|" then
    trimmed = trimmed:sub(2)
  end
  if trimmed:sub(-1) == "|" and trimmed:sub(-2) ~= "\\|" then
    trimmed = trimmed:sub(1, -2)
  end

  local cells = {}
  local cur = {}
  local i = 1
  local len = #trimmed
  local in_code = false
  while i <= len do
    local ch = trimmed:sub(i, i)
    if ch == "\\" and i < len and trimmed:sub(i + 1, i + 1) == "|" then
      cur[#cur + 1] = "|"
      i = i + 2
    elseif ch == "`" then
      in_code = not in_code
      cur[#cur + 1] = ch
      i = i + 1
    elseif ch == "|" and not in_code then
      local cell_str = table.concat(cur):gsub("^%s+", ""):gsub("%s+$", "")
      cells[#cells + 1] = cell_str
      cur = {}
      i = i + 1
    else
      cur[#cur + 1] = ch
      i = i + 1
    end
  end
  local last_str = table.concat(cur):gsub("^%s+", ""):gsub("%s+$", "")
  cells[#cells + 1] = last_str
  return cells
end

local function is_pipe_separator_line(line)
  if not line then return false, nil end
  local trimmed = line:gsub("^%s+", ""):gsub("%s+$", "")
  if not trimmed:find("|", 1, true) or not trimmed:find("-", 1, true) then
    return false, nil
  end
  local cells = split_pipe_row(trimmed)
  if #cells == 0 then
    return false, nil
  end
  local alignments = {}
  for idx, cell in ipairs(cells) do
    if not cell:match("^:?%-+:?$") then
      return false, nil
    end
    local left_colon = cell:sub(1, 1) == ":"
    local right_colon = cell:sub(-1) == ":"
    if left_colon and right_colon then
      alignments[idx] = "center"
    elseif right_colon then
      alignments[idx] = "right"
    elseif left_colon then
      alignments[idx] = "left"
    else
      alignments[idx] = "left"
    end
  end
  return true, alignments
end

local function is_hr_line(line)
  local s = line:gsub("^%s+", ""):gsub("%s+$", "")
  if #s < 3 then return false end
  if s:match("^%-[%-%s]+$") and select(2, s:gsub("%-", "")) >= 3 then
    return true
  end
  if s:match("^%*[%*%s]+$") and select(2, s:gsub("%*", "")) >= 3 then
    return true
  end
  if s:match("^_[_%s]+$") and select(2, s:gsub("_", "")) >= 3 then
    return true
  end
  return false
end

local function match_code_fence(line)
  local ticks, lang = line:match("^%s*(```+)%s*(.-)%s*$")
  if ticks then
    return ticks, lang
  end
  local tildes, tlang = line:match("^%s*(~~~+)%s*(.-)%s*$")
  if tildes then
    return tildes, tlang
  end
  return nil, nil
end

local function match_atx_heading(line)
  local hashes, rest = line:match("^%s*(#+)%s+(.-)%s*$")
  if not hashes then
    hashes = line:match("^%s*(#+)%s*$")
    rest = ""
  end
  if hashes and #hashes >= 1 and #hashes <= 6 then
    rest = rest:gsub("%s+#+%s*$", "")
    return #hashes, rest
  end
  return nil, nil
end

function M.markdown_to_ast(lines, opts)
  opts = opts or {}
  local doc = {
    metadata = {},
    blocks = {},
    footnotes = {},
  }

  local n = #(lines or {})
  local idx = 1

  -- 1. Check optional YAML frontmatter at top of document
  if idx <= n and lines[1] == "---" then
    local close_idx = nil
    local valid_yaml = true
    local has_known_key = false
    local temp_meta = {}
    for j = 2, math.min(n, 30) do
      if lines[j] == "---" or lines[j] == "..." then
        close_idx = j
        break
      end
      local k, v = lines[j]:match("^([%w_%-]+):%s*(.*)$")
      if k then
        local kl = k:lower()
        v = v:gsub('^"(.*)"$', "%1"):gsub("^'(.*)'$", "%1")
        temp_meta[kl] = v
        if kl == "title" or kl == "author" or kl == "creator" or kl == "subject" or kl == "keywords" or kl == "description" or kl == "date" then
          has_known_key = true
        end
      else
        valid_yaml = false
        break
      end
    end
    if close_idx and valid_yaml and has_known_key then
      doc.metadata = temp_meta
      idx = close_idx + 1
    end
  end

  local fn_map = {}
  -- Indentation stack for list nesting (supports 2-space, 3-space, or 4-space indentation)
  local list_indent_stack = { [0] = 0 }
  local cur_list_level = 0

  local function resolve_indent_level(indent_str)
    local cols = #(indent_str:gsub("\t", "    "))
    local prev_block = doc.blocks[#doc.blocks]
    if not prev_block or prev_block.type ~= "list_item" then
      list_indent_stack = { [0] = 0 }
      cur_list_level = 0
    end
    if cols <= 0 then
      cur_list_level = 0
      list_indent_stack = { [0] = 0 }
      return 0
    end
    if cols > (list_indent_stack[cur_list_level] or 0) then
      cur_list_level = math.min(8, cur_list_level + 1)
      list_indent_stack[cur_list_level] = cols
    else
      while cur_list_level > 0 and cols < (list_indent_stack[cur_list_level] or 0) do
        list_indent_stack[cur_list_level] = nil
        cur_list_level = cur_list_level - 1
      end
    end
    return cur_list_level
  end

  -- ::: {attrs} ... ::: sections; their attributes apply to every block inside
  local div_stack = {}
  local function current_div_attrs()
    if #div_stack == 0 then
      return nil
    end
    local merged = { kv = {}, classes = {} }
    for _, a in ipairs(div_stack) do
      for k, v in pairs(a.kv) do
        merged.kv[k] = v
      end
      for k, v in pairs(a.classes) do
        merged.classes[k] = v
      end
    end
    return merged
  end

  while idx <= n do
    local block_start_idx, block_count_before = idx, #doc.blocks
    local line = lines[idx]
    local fence, fence_lang = match_code_fence(line)
    local h_level, h_text = match_atx_heading(line)
    local div_open = line:match("^%s*:::+%s*{(.*)}%s*$")
    local div_class = not div_open and line:match("^%s*:::+%s*([%w%-_]+)%s*$")

    if div_open or div_class then
      div_stack[#div_stack + 1] = div_open and (M.parse_attrs(div_open) or { kv = {}, classes = {} })
        or { kv = { style = div_class }, classes = {} }
      idx = idx + 1

    elseif line:match("^%s*:::+%s*$") and #div_stack > 0 then
      table.remove(div_stack)
      idx = idx + 1

    -- Page break
    elseif line:match("^%s*\\newpage%s*$") or line:match("^%s*\\pagebreak%s*$") then
      doc.blocks[#doc.blocks + 1] = { type = "pagebreak" }
      idx = idx + 1

    -- Skip blank lines
    elseif line:match("^%s*$") then
      idx = idx + 1

    -- Fenced code block ``` or ~~~
    elseif fence then
      local fence_char = fence:sub(1, 1)
      local fence_len = #fence
      idx = idx + 1
      local code_lines = {}
      while idx <= n do
        local cl = lines[idx]
        local closing = cl:match("^%s*(" .. fence_char .. "+)%s*$")
        if closing and #closing >= fence_len then
          idx = idx + 1
          break
        end
        code_lines[#code_lines + 1] = cl
        idx = idx + 1
      end
      doc.blocks[#doc.blocks + 1] = {
        type = "code_block",
        lang = fence_lang or "",
        text = table.concat(code_lines, "\n"),
      }

    -- Display math block $$ ... $$
    elseif line:match("^%s*%$%$(.*)$") then
      local rest = line:match("^%s*%$%$(.*)$")
      if rest:match("^(.-)%$%$%s*$") and rest:gsub("%$%$%s*$", "") ~= "" then
        local single_math = rest:match("^(.-)%$%$%s*$")
        doc.blocks[#doc.blocks + 1] = {
          type = "math_block",
          text = single_math:gsub("^%s+", ""):gsub("%s+$", ""),
        }
        idx = idx + 1
      else
        local math_lines = {}
        if rest:find("%S") then
          math_lines[#math_lines + 1] = rest
        end
        idx = idx + 1
        while idx <= n do
          local ml = lines[idx]
          if ml:match("^(.-)%$%$%s*$") then
            local before_close = ml:match("^(.-)%$%$%s*$")
            if before_close:find("%S") then
              math_lines[#math_lines + 1] = before_close
            end
            idx = idx + 1
            break
          end
          math_lines[#math_lines + 1] = ml
          idx = idx + 1
        end
        doc.blocks[#doc.blocks + 1] = {
          type = "math_block",
          text = table.concat(math_lines, "\n"),
        }
      end

    -- Footnote definition [^id]: text
    elseif line:match("^%[%^([^%]%s]+)%]:%s*(.*)$") then
      local fn_id, first_txt = line:match("^%[%^([^%]%s]+)%]:%s*(.*)$")
      local fn_lines = { first_txt }
      idx = idx + 1
      while idx <= n do
        local nl = lines[idx]
        if nl:match("^    (.*)$") or nl:match("^\t(.*)$") then
          fn_lines[#fn_lines + 1] = nl:gsub("^%s+", "")
          idx = idx + 1
        else
          break
        end
      end
      local fn_inlines = M.parse_inlines(table.concat(fn_lines, " "))
      if not fn_map[fn_id] then
        local entry = { id = fn_id, inlines = fn_inlines }
        fn_map[fn_id] = entry
        doc.footnotes[#doc.footnotes + 1] = entry
      else
        fn_map[fn_id].inlines = fn_inlines
      end

    -- ATX Heading (# .. ######)
    elseif h_level then
      doc.blocks[#doc.blocks + 1] = {
        type = "heading",
        level = h_level,
        inlines = M.parse_inlines(h_text),
      }
      idx = idx + 1

    -- Pipe Table (header line + separator line)
    elseif line:find("|", 1, true) and idx < n and is_pipe_separator_line(lines[idx + 1]) then
      local _, alignments = is_pipe_separator_line(lines[idx + 1])
      local raw_headers = split_pipe_row(line)
      local num_cols = math.max(#raw_headers, #alignments)
      local headers = {}
      for c = 1, num_cols do
        headers[c] = M.parse_inlines(raw_headers[c] or "")
        if not alignments[c] then
          alignments[c] = "left"
        end
      end

      idx = idx + 2
      local rows = {}
      while idx <= n do
        local rline = lines[idx]
        if rline:match("^%s*$") or not rline:find("|", 1, true) then
          break
        end
        local raw_cells = split_pipe_row(rline)
        local row = {}
        for c = 1, num_cols do
          row[c] = M.parse_inlines(raw_cells[c] or "")
        end
        rows[#rows + 1] = row
        idx = idx + 1
      end

      doc.blocks[#doc.blocks + 1] = {
        type = "table",
        alignments = alignments,
        headers = headers,
        rows = rows,
      }

    -- Horizontal Rule (---, ***, ___)
    elseif is_hr_line(line) then
      doc.blocks[#doc.blocks + 1] = { type = "hr" }
      idx = idx + 1

    -- Blockquote (> ...)
    elseif line:match("^%s*>") then
      local first_prefix = line:match("^%s*(>+)")
      local bq_level = first_prefix and #first_prefix or 1
      local bq_lines = {}
      while idx <= n and lines[idx]:match("^%s*>") do
        local prefix, rest = lines[idx]:match("^%s*(>+)%s?(.*)$")
        local p_lvl = prefix and #prefix or 1
        if rest:match("^%s*$") then
          idx = idx + 1
          break
        end
        if #bq_lines > 0 and p_lvl ~= bq_level then
          break
        end
        bq_lines[#bq_lines + 1] = rest or ""
        idx = idx + 1
      end
      if #bq_lines > 0 then
        doc.blocks[#doc.blocks + 1] = {
          type = "blockquote",
          level = bq_level,
          inlines = M.parse_inlines(table.concat(bq_lines, " ")),
        }
      end

    -- Unordered list item (-, *, +)
    elseif line:match("^(%s*)[%-%*%+]%s+(.*)$") then
      local indent, item_text = line:match("^(%s*)[%-%*%+]%s+(.*)$")
      local level = resolve_indent_level(indent)
      local checked = nil
      local box, after_box = item_text:match("^%[([ xX])%]%s+(.*)$")
      if box then
        checked = (box:lower() == "x")
        item_text = after_box
      end
      local item_parts = { item_text }
      idx = idx + 1
      while item_parts[#item_parts]:match("\\%s*$") and idx <= n and not lines[idx]:match("^%s*$") do
        item_parts[#item_parts] = item_parts[#item_parts]:gsub("\\%s*$", "")
        item_parts[#item_parts + 1] = lines[idx]
        idx = idx + 1
      end
      local item_inlines = {}
      for pi, part in ipairs(item_parts) do
        for _, inl in ipairs(M.parse_inlines(part)) do
          item_inlines[#item_inlines + 1] = inl
        end
        if pi < #item_parts then
          item_inlines[#item_inlines + 1] = { type = "linebreak" }
        end
      end
      doc.blocks[#doc.blocks + 1] = {
        type = "list_item",
        ordered = false,
        level = level,
        checked = checked,
        inlines = M.merge_inlines(item_inlines),
      }

    -- Ordered list item (1. or 1))
    elseif line:match("^(%s*)(%d+)[%.%)]%s+(.*)$") then
      local indent, num_str, item_text = line:match("^(%s*)(%d+)[%.%)]%s+(.*)$")
      local level = resolve_indent_level(indent)
      local item_parts = { item_text }
      idx = idx + 1
      while item_parts[#item_parts]:match("\\%s*$") and idx <= n and not lines[idx]:match("^%s*$") do
        item_parts[#item_parts] = item_parts[#item_parts]:gsub("\\%s*$", "")
        item_parts[#item_parts + 1] = lines[idx]
        idx = idx + 1
      end
      local item_inlines = {}
      for pi, part in ipairs(item_parts) do
        for _, inl in ipairs(M.parse_inlines(part)) do
          item_inlines[#item_inlines + 1] = inl
        end
        if pi < #item_parts then
          item_inlines[#item_inlines + 1] = { type = "linebreak" }
        end
      end
      doc.blocks[#doc.blocks + 1] = {
        type = "list_item",
        ordered = true,
        level = level,
        number = tonumber(num_str) or 1,
        inlines = M.merge_inlines(item_inlines),
      }

    -- Setext Heading (Paragraph line followed by === or ---)
    elseif idx < n and not line:match("^%s*$") and (lines[idx + 1]:match("^%s*===+%s*$") or lines[idx + 1]:match("^%s*%-%-%-+%s*$")) then
      local lvl = lines[idx + 1]:find("=", 1, true) and 1 or 2
      doc.blocks[#doc.blocks + 1] = {
        type = "heading",
        level = lvl,
        inlines = M.parse_inlines(line:gsub("^%s+", ""):gsub("%s+$", "")),
      }
      idx = idx + 2

    -- Normal Paragraph
    else
      local para_parts = {}
      while idx <= n do
        local pl = lines[idx]
        if pl:match("^%s*$")
          or match_code_fence(pl) ~= nil
          or match_atx_heading(pl) ~= nil
          or pl:match("^%s*%$%$")
          or pl:match("^%s*>")
          or pl:match("^(%s*)[%-%*%+]%s+")
          or pl:match("^(%s*)(%d+)[%.%)]%s+")
          or pl:match("^%[%^([^%]%s]+)%]:%s*")
          or is_hr_line(pl)
          or (pl:find("|", 1, true) and idx < n and is_pipe_separator_line(lines[idx + 1]))
        then
          break
        end

        para_parts[#para_parts + 1] = pl
        idx = idx + 1

        -- If this line is a standalone image, keep it as its own paragraph
        if pl:match("^%s*!%[.-%]%(.-%)%s*(%b{})%s*$") or pl:match("^%s*!%[.-%]%(.-%)%s*$") then
          break
        end
      end

      if #para_parts > 0 then
        -- Combine paragraph lines respecting hard line breaks (\ at EOL or 2+ trailing spaces)
        local combined_inlines = {}
        for i, ptext in ipairs(para_parts) do
          local hard_break = false
          if ptext:match("\\%s*$") then
            hard_break = true
            ptext = ptext:gsub("\\%s*$", "")
          elseif ptext:match("  +$") then
            hard_break = true
            ptext = ptext:gsub("  +$", "")
          elseif opts.wrap == "preserve" and i < #para_parts then
            hard_break = true
          end

          local line_inlines = M.parse_inlines(ptext)
          for _, inl in ipairs(line_inlines) do
            combined_inlines[#combined_inlines + 1] = inl
          end

          if i < #para_parts then
            if hard_break then
              combined_inlines[#combined_inlines + 1] = { type = "linebreak" }
            else
              combined_inlines[#combined_inlines + 1] = { type = "text", text = " " }
            end
          end
        end

        doc.blocks[#doc.blocks + 1] = {
          type = "paragraph",
          inlines = M.merge_inlines(combined_inlines),
        }
      else
        idx = idx + 1
      end
    end

    -- Record the source line range of each block (used to diff against the loaded document)
    if opts.block_ranges then
      for k = block_count_before + 1, #doc.blocks do
        opts.block_ranges[k] = { block_start_idx, idx - 1 }
      end
    end
    local div_attrs = current_div_attrs()
    if div_attrs then
      for k = block_count_before + 1, #doc.blocks do
        doc.blocks[k].attrs = div_attrs
      end
    end
  end

  return doc
end

--------------------------------------------------------------------------------
-- Block Renderer (Document AST -> Markdown lines)
--------------------------------------------------------------------------------

local function wrap_text_line(str, width)
  width = width or 72
  if #str <= width then
    return { str }
  end
  local result = {}
  local current = ""
  for word in str:gmatch("%S+") do
    if current == "" then
      current = word
    elseif #current + 1 + #word <= width then
      current = current .. " " .. word
    else
      result[#result + 1] = current
      current = word
    end
  end
  if current ~= "" then
    result[#result + 1] = current
  end
  return #result > 0 and result or { "" }
end

-- Canonical "{...}" text of block attributes (nil when none)
local BLOCK_ATTR_ORDER = { "style", "align", "font", "size", "color", "highlight" }
function M.block_attrs_sig(attrs)
  if not attrs then
    return nil
  end
  local parts = {}
  for _, k in ipairs(BLOCK_ATTR_ORDER) do
    local v = attrs.kv and attrs.kv[k]
    if v then
      parts[#parts + 1] = (v:find("[%s\"]") or k == "font" or k == "style") and string.format('%s="%s"', k, v)
        or (k .. "=" .. v)
    end
  end
  local classes = {}
  for c in pairs(attrs.classes or {}) do
    classes[#classes + 1] = "." .. c
  end
  table.sort(classes)
  vim.list_extend(parts, classes)
  if #parts == 0 then
    return nil
  end
  return "{" .. table.concat(parts, " ") .. "}"
end

function M.ast_to_markdown_lines(doc, opts)
  opts = opts or {}
  local lines = {}
  local blocks = (doc and doc.blocks) or {}
  local wrap_mode = opts.wrap or "none"

  local function append_str_lines(str)
    for l in (str .. "\n"):gmatch("(.-)\n") do
      lines[#lines + 1] = l
    end
  end

  -- Optional YAML frontmatter
  if opts.frontmatter and doc and doc.metadata and next(doc.metadata) ~= nil then
    lines[#lines + 1] = "---"
    local keys = { "title", "author", "subject", "keywords", "description" }
    for _, k in ipairs(keys) do
      if doc.metadata[k] and doc.metadata[k] ~= "" then
        lines[#lines + 1] = string.format("%s: %s", k, doc.metadata[k])
      end
    end
    lines[#lines + 1] = "---"
    lines[#lines + 1] = ""
  end

  local list_counters = {}
  local rendered_list_level = -1
  local open_div = nil

  for i, block in ipairs(blocks) do
    local btype = block.type
    local prev_block = blocks[i - 1]
    local prev_type = prev_block and prev_block.type or nil

    -- Reset ordered list counters when leaving a list
    if btype ~= "list_item" then
      list_counters = {}
    end

    -- Add blank separator line between blocks (except consecutive list items of compatible kind)
    if i > 1 then
      local same_list = false
      if btype == "list_item" and prev_type == "list_item" then
        if (block.level or 0) > 0 or (prev_block.level or 0) > 0 or (block.ordered == prev_block.ordered) then
          same_list = true
        end
      end
      if not same_list then
        lines[#lines + 1] = ""
      end
    end

    -- Open/close ::: sections when the block attributes change
    local sig = M.block_attrs_sig(block.attrs)
    if sig ~= open_div then
      if open_div then
        if lines[#lines] ~= "" then
          lines[#lines + 1] = ""
        end
        lines[#lines + 1] = ":::"
        lines[#lines + 1] = ""
      end
      if sig then
        if #lines > 0 and lines[#lines] ~= "" then
          lines[#lines + 1] = ""
        end
        lines[#lines + 1] = "::: " .. sig
        lines[#lines + 1] = ""
      end
      open_div = sig
    end

    local first_line = #lines + 1

    if btype == "pagebreak" then
      lines[#lines + 1] = "\\newpage"

    elseif btype == "heading" then
      local lvl = math.min(math.max(tonumber(block.level) or 1, 1), 6)
      local prefix = string.rep("#", lvl) .. " "
      -- A heading is a single Markdown line; line breaks inside it become <br>
      local text = M.render_inlines(block.inlines or {}, vim.tbl_extend("force", opts, { single_line = true }))
      text = text:gsub("\n", " ")
      append_str_lines(prefix .. text)

    elseif btype == "paragraph" then
      local text = M.render_inlines(block.inlines or {}, opts)
      local esc_lines = vim.split(text, "\n", { plain = true })
      for k, l in ipairs(esc_lines) do
        esc_lines[k] = escape_block_start(l)
      end
      text = table.concat(esc_lines, "\n")
      if wrap_mode == "auto" then
        for _, subline in ipairs(vim.split(text, "\n", { plain = true })) do
          for _, wl in ipairs(wrap_text_line(subline, 72)) do
            lines[#lines + 1] = wl
          end
        end
      else
        append_str_lines(text)
      end

    elseif btype == "list_item" then
      local lvl = math.min(math.max(tonumber(block.level) or 0, 0), 8)
      -- Markdown nesting can only go one level deeper per item (Word lists may skip levels)
      local prev_lvl = (prev_type == "list_item") and rendered_list_level or -1
      lvl = math.min(lvl, prev_lvl + 1)
      rendered_list_level = lvl
      -- Clear deeper counters
      for k, _ in pairs(list_counters) do
        if k > lvl then list_counters[k] = nil end
      end
      local indent = string.rep("  ", lvl)
      local marker
      if block.ordered then
        local n_val = block.number or ((list_counters[lvl] or 0) + 1)
        list_counters[lvl] = n_val
        marker = tostring(n_val) .. ". "
      else
        list_counters[lvl] = nil
        marker = "- "
      end
      local checkbox = ""
      if block.checked == true then
        checkbox = "[x] "
      elseif block.checked == false then
        checkbox = "[ ] "
      end
      local text = M.render_inlines(block.inlines or {}, opts)
      append_str_lines(indent .. marker .. checkbox .. text)

    elseif btype == "blockquote" then
      local lvl = math.max(tonumber(block.level) or 1, 1)
      local prefix = string.rep("> ", lvl)
      local text = M.render_inlines(block.inlines or {}, opts)
      for _, subline in ipairs(vim.split(text, "\n", { plain = true })) do
        lines[#lines + 1] = prefix .. subline
      end

    elseif btype == "code_block" then
      local fence = "```"
      if (block.text or ""):find("```", 1, true) then
        fence = "~~~~"
      end
      lines[#lines + 1] = fence .. (block.lang or "")
      append_str_lines(block.text or "")
      lines[#lines + 1] = fence

    elseif btype == "math_block" then
      lines[#lines + 1] = "$$"
      append_str_lines(block.text or "")
      lines[#lines + 1] = "$$"

    elseif btype == "hr" then
      lines[#lines + 1] = "---"

    elseif btype == "table" then
      local tbl_opts = vim.tbl_extend("force", opts, { in_table = true })
      local headers = block.headers or {}
      local rows = block.rows or {}
      local alignments = block.alignments or {}
      local num_cols = #headers
      for _, r in ipairs(rows) do
        if #r > num_cols then num_cols = #r end
      end
      if num_cols == 0 then num_cols = 1 end

      local hdr_cells = {}
      local sep_cells = {}
      for c = 1, num_cols do
        hdr_cells[c] = M.render_inlines(headers[c] or {}, tbl_opts)
        local al = alignments[c] or "left"
        if al == "center" then
          sep_cells[c] = ":---:"
        elseif al == "right" then
          sep_cells[c] = "---:"
        else
          sep_cells[c] = ":---"
        end
      end
      lines[#lines + 1] = "| " .. table.concat(hdr_cells, " | ") .. " |"
      lines[#lines + 1] = "| " .. table.concat(sep_cells, " | ") .. " |"

      for _, r in ipairs(rows) do
        local row_cells = {}
        for c = 1, num_cols do
          row_cells[c] = M.render_inlines(r[c] or {}, tbl_opts)
        end
        lines[#lines + 1] = "| " .. table.concat(row_cells, " | ") .. " |"
      end
    end

    if opts.block_ranges then
      opts.block_ranges[i] = { first_line, #lines }
    end
  end

  if open_div then
    lines[#lines + 1] = ""
    lines[#lines + 1] = ":::"
  end

  if doc and doc.footnotes and #doc.footnotes > 0 then
    if #lines > 0 then
      lines[#lines + 1] = ""
    end
    for _, fn in ipairs(doc.footnotes) do
      local fn_text = M.render_inlines(fn.inlines or {}, vim.tbl_extend("force", opts, { single_line = true }))
      fn_text = fn_text:gsub("\n", " ")
      lines[#lines + 1] = string.format("[^%s]: %s", tostring(fn.id), fn_text)
    end
  end

  if #lines == 0 then
    lines[1] = ""
  end

  return lines
end

function M.ast_to_markdown(doc, opts)
  return table.concat(M.ast_to_markdown_lines(doc, opts), "\n")
end

return M
