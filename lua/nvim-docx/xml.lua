-- Pure-Lua XML parser, DOM helper, and serializer for OOXML (.docx) documents.

local M = {}

local function codepoint_to_utf8(cp)
  if not cp or cp < 0 then
    return ""
  end
  if cp <= 0x7F then
    return string.char(cp)
  elseif cp <= 0x7FF then
    local b1 = 0xC0 + math.floor(cp / 64)
    local b2 = 0x80 + (cp % 64)
    return string.char(b1, b2)
  elseif cp <= 0xFFFF then
    local b1 = 0xE0 + math.floor(cp / 4096)
    local b2 = 0x80 + (math.floor(cp / 64) % 64)
    local b3 = 0x80 + (cp % 64)
    return string.char(b1, b2, b3)
  elseif cp <= 0x10FFFF then
    local b1 = 0xF0 + math.floor(cp / 262144)
    local b2 = 0x80 + (math.floor(cp / 4096) % 64)
    local b3 = 0x80 + (math.floor(cp / 64) % 64)
    local b4 = 0x80 + (cp % 64)
    return string.char(b1, b2, b3, b4)
  end
  return ""
end

function M.unescape(str)
  if not str or str == "" or not str:find("&", 1, true) then
    return str or ""
  end
  str = str:gsub("&#x([0-9a-fA-F]+);", function(hex)
    return codepoint_to_utf8(tonumber(hex, 16))
  end)
  str = str:gsub("&#([0-9]+);", function(dec)
    return codepoint_to_utf8(tonumber(dec, 10))
  end)
  str = str:gsub("&lt;", "<")
  str = str:gsub("&gt;", ">")
  str = str:gsub("&quot;", '"')
  str = str:gsub("&apos;", "'")
  str = str:gsub("&amp;", "&")
  return str
end

function M.escape_text(str)
  if not str or str == "" then
    return ""
  end
  -- Strip invalid XML 1.0 control chars (keep \t \n \r)
  str = str:gsub("[%z\1-\8\11\12\14-\31]", "")
  str = str:gsub("&", "&amp;")
  str = str:gsub("<", "&lt;")
  str = str:gsub(">", "&gt;")
  return str
end

function M.escape_attr(str)
  if not str or str == "" then
    return ""
  end
  str = M.escape_text(tostring(str))
  str = str:gsub('"', "&quot;")
  return str
end

local function local_tag_name(tag)
  local colon = tag:find(":", 1, true)
  if colon then
    return tag:sub(colon + 1)
  end
  return tag
end

function M.elem(tag, attrs, children)
  local attr_map = {}
  local attr_order = {}
  if attrs then
    -- Support ordered list entries `{ { "k", "v" }, ... }` or map `{ k = "v" }`
    for k, v in pairs(attrs) do
      if type(k) == "number" and type(v) == "table" then
        local ak, av = v[1], v[2]
        if ak and av ~= nil then
          if attr_map[ak] == nil then
            attr_order[#attr_order + 1] = ak
          end
          attr_map[ak] = tostring(av)
        end
      elseif type(k) == "string" then
        attr_map[k] = tostring(v)
        attr_order[#attr_order + 1] = k
      end
    end
    table.sort(attr_order)
  end
  return {
    tag = tag,
    local_name = local_tag_name(tag),
    attrs = attr_map,
    attr_order = attr_order,
    children = children or {},
  }
end

function M.text_node(text)
  return { text = tostring(text or "") }
end

function M.set_attr(node, name, val)
  if not node.attrs then
    node.attrs = {}
  end
  if not node.attr_order then
    node.attr_order = {}
  end
  if val == nil then
    node.attrs[name] = nil
    for i = #node.attr_order, 1, -1 do
      if node.attr_order[i] == name then
        table.remove(node.attr_order, i)
      end
    end
  else
    if node.attrs[name] == nil then
      node.attr_order[#node.attr_order + 1] = name
    end
    node.attrs[name] = tostring(val)
  end
end

function M.get_attr(node, name)
  if not node or not node.attrs then
    return nil
  end
  if node.attrs[name] ~= nil then
    return node.attrs[name]
  end
  local target_local = local_tag_name(name)
  for k, v in pairs(node.attrs) do
    if local_tag_name(k) == target_local then
      return v
    end
  end
  return nil
end

local function parse_attributes(attr_str)
  local attrs = {}
  local attr_order = {}
  if not attr_str or attr_str == "" then
    return attrs, attr_order
  end
  local pos = 1
  local len = #attr_str
  while pos <= len do
    local s, e, name, quote, val = attr_str:find("([%w_:%-%.]+)%s*=%s*([\"'])(.-)%2", pos)
    if not s then
      break
    end
    local decoded = M.unescape(val)
    if attrs[name] == nil then
      attr_order[#attr_order + 1] = name
    end
    attrs[name] = decoded
    pos = e + 1
  end
  return attrs, attr_order
end

function M.parse(xml_str)
  if not xml_str or xml_str == "" then
    return nil, "Empty XML input"
  end

  -- Strip UTF-8 BOM if present
  if xml_str:sub(1, 3) == "\239\187\191" then
    xml_str = xml_str:sub(4)
  end

  local root = { tag = "#document", local_name = "#document", attrs = {}, attr_order = {}, children = {} }
  local stack = { root }
  local pos = 1
  local len = #xml_str

  while pos <= len do
    local lt = xml_str:find("<", pos, true)
    if not lt then
      local tail = xml_str:sub(pos)
      if tail ~= "" then
        local top = stack[#stack]
        top.children[#top.children + 1] = { text = M.unescape(tail) }
      end
      break
    end

    if lt > pos then
      local raw_text = xml_str:sub(pos, lt - 1)
      local top = stack[#stack]
      -- Preserve all text inside leaf text elements; ignore pure formatting whitespace outside
      local lname = top.local_name
      if lname == "t" or lname == "delText" or lname == "instrText" or raw_text:find("%S") then
        top.children[#top.children + 1] = { text = M.unescape(raw_text) }
      end
    end

    local next_ch = xml_str:sub(lt + 1, lt + 1)
    if next_ch == "?" then
      -- Processing instruction <? ... ?>
      local close_pi = xml_str:find("?>", lt + 2, true)
      if not close_pi then
        return nil, "Unclosed XML processing instruction"
      end
      pos = close_pi + 2
    elseif next_ch == "!" then
      if xml_str:sub(lt + 2, lt + 3) == "--" then
        -- Comment <!-- ... -->
        local close_comment = xml_str:find("-->", lt + 4, true)
        if not close_comment then
          return nil, "Unclosed XML comment"
        end
        pos = close_comment + 3
      elseif xml_str:sub(lt + 2, lt + 8) == "[CDATA[" then
        -- CDATA <![CDATA[ ... ]]>
        local close_cdata = xml_str:find("]]>", lt + 9, true)
        if not close_cdata then
          return nil, "Unclosed CDATA section"
        end
        local cdata_text = xml_str:sub(lt + 9, close_cdata - 1)
        local top = stack[#stack]
        top.children[#top.children + 1] = { text = cdata_text }
        pos = close_cdata + 3
      else
        -- DOCTYPE or other declaration (handle optional [...] internal subset)
        local scan_decl = lt + 2
        local bracket_depth = 0
        local gt = nil
        while scan_decl <= len do
          local ch = xml_str:sub(scan_decl, scan_decl)
          if ch == "[" then
            bracket_depth = bracket_depth + 1
          elseif ch == "]" and bracket_depth > 0 then
            bracket_depth = bracket_depth - 1
          elseif ch == ">" and bracket_depth == 0 then
            gt = scan_decl
            break
          end
          scan_decl = scan_decl + 1
        end
        if not gt then
          return nil, "Unclosed XML declaration"
        end
        pos = gt + 1
      end
    elseif next_ch == "/" then
      -- Closing tag </tag>
      local gt = xml_str:find(">", lt + 2, true)
      if not gt then
        return nil, "Unclosed XML closing tag"
      end
      if #stack > 1 then
        table.remove(stack)
      end
      pos = gt + 1
    else
      -- Opening or self-closing tag
      -- Scan to matching '>' taking quotes into account
      local scan = lt + 1
      local in_quote = nil
      local gt = nil
      while scan <= len do
        local ch = xml_str:sub(scan, scan)
        if in_quote then
          if ch == in_quote then
            in_quote = nil
          end
        else
          if ch == '"' or ch == "'" then
            in_quote = ch
          elseif ch == ">" then
            gt = scan
            break
          end
        end
        scan = scan + 1
      end

      if not gt then
        return nil, "Unclosed XML element tag"
      end

      local tag_body = xml_str:sub(lt + 1, gt - 1)
      local self_closing = false
      if tag_body:sub(-1) == "/" then
        self_closing = true
        tag_body = tag_body:sub(1, -2)
      end

      local tag_name, attr_str = tag_body:match("^%s*([%w_:%-%.]+)(.*)$")
      if not tag_name then
        return nil, "Malformed XML tag: <" .. tag_body .. ">"
      end

      local attrs, attr_order = parse_attributes(attr_str)
      local node = {
        tag = tag_name,
        local_name = local_tag_name(tag_name),
        attrs = attrs,
        attr_order = attr_order,
        children = {},
      }

      local top = stack[#stack]
      top.children[#top.children + 1] = node
      if not self_closing then
        stack[#stack + 1] = node
      end

      pos = gt + 1
    end
  end

  for _, child in ipairs(root.children) do
    if child.tag then
      return child
    end
  end

  return nil, "No root XML element found"
end

function M.matches(node, tag_or_local)
  if not node or not node.tag then
    return false
  end
  if not tag_or_local then
    return true
  end
  return node.tag == tag_or_local or node.local_name == tag_or_local
end

function M.find_child(node, tag_or_local)
  if not node or not node.children then
    return nil
  end
  for _, child in ipairs(node.children) do
    if M.matches(child, tag_or_local) then
      return child
    end
  end
  return nil
end

function M.find_children(node, tag_or_local)
  local res = {}
  if not node or not node.children then
    return res
  end
  for _, child in ipairs(node.children) do
    if M.matches(child, tag_or_local) then
      res[#res + 1] = child
    end
  end
  return res
end

-- Pre-order walk over element descendants with an explicit stack (documents can nest tables
-- thousands of levels deep, which overflows the Lua stack when recursing). Returning true from fn
-- stops the walk.
function M.walk(node, fn)
  local stack = { { node, 1 } }
  while #stack > 0 do
    local top = stack[#stack]
    local children = top[1].children
    local idx = top[2]
    if not children or idx > #children then
      stack[#stack] = nil
    else
      top[2] = idx + 1
      local child = children[idx]
      if child.tag then
        if fn(child) then
          return true
        end
        stack[#stack + 1] = { child, 1 }
      end
    end
  end
  return false
end

function M.find_all(node, tag_or_local, acc)
  acc = acc or {}
  if not node or not node.children then
    return acc
  end
  M.walk(node, function(child)
    if M.matches(child, tag_or_local) then
      acc[#acc + 1] = child
    end
  end)
  return acc
end

-- Deep copy of an element tree without recursion
function M.clone(node)
  local function shallow(n)
    if n.text ~= nil then
      return { text = n.text }
    end
    local attrs, order = {}, {}
    for k, v in pairs(n.attrs or {}) do
      attrs[k] = v
    end
    for i, k in ipairs(n.attr_order or {}) do
      order[i] = k
    end
    return { tag = n.tag, local_name = n.local_name, attrs = attrs, attr_order = order, children = {} }
  end
  local root = shallow(node)
  local stack = { { node, root } }
  while #stack > 0 do
    local pair = table.remove(stack)
    for _, c in ipairs(pair[1].children or {}) do
      local copy = shallow(c)
      pair[2].children[#pair[2].children + 1] = copy
      if c.children then
        stack[#stack + 1] = { c, copy }
      end
    end
  end
  return root
end

function M.get_direct_text(node)
  if not node or not node.children then
    return ""
  end
  local parts = {}
  for _, child in ipairs(node.children) do
    if child.text then
      parts[#parts + 1] = child.text
    end
  end
  return table.concat(parts)
end

function M.get_all_text(node, acc)
  acc = acc or {}
  if not node then
    return ""
  end
  if node.text then
    acc[#acc + 1] = node.text
    return table.concat(acc)
  end
  for _, child in ipairs(node.children or {}) do
    if child.text then
      acc[#acc + 1] = child.text
    elseif child.tag then
      M.get_all_text(child, acc)
    end
  end
  return table.concat(acc)
end

local function serialize_open_tag(node, out)
  out[#out + 1] = "<"
  out[#out + 1] = node.tag

  if node.attrs then
    local order = node.attr_order or {}
    local emitted = {}
    for _, k in ipairs(order) do
      local v = node.attrs[k]
      if v ~= nil and not emitted[k] then
        emitted[k] = true
        out[#out + 1] = " "
        out[#out + 1] = k
        out[#out + 1] = '="'
        out[#out + 1] = M.escape_attr(v)
        out[#out + 1] = '"'
      end
    end
    local extra = {}
    for k, _ in pairs(node.attrs) do
      if not emitted[k] then
        extra[#extra + 1] = k
      end
    end
    if #extra > 0 then
      table.sort(extra)
      for _, k in ipairs(extra) do
        out[#out + 1] = " "
        out[#out + 1] = k
        out[#out + 1] = '="'
        out[#out + 1] = M.escape_attr(node.attrs[k])
        out[#out + 1] = '"'
      end
    end
  end
end

-- Iterative serializer (see M.walk for why recursion is avoided)
local function serialize_node(root, out)
  local stack = { { root, 0 } }
  while #stack > 0 do
    local top = stack[#stack]
    local node, idx = top[1], top[2]
    if idx == 0 then
      if node.text ~= nil then
        out[#out + 1] = M.escape_text(node.text)
        stack[#stack] = nil
      elseif not node.tag then
        stack[#stack] = nil
      else
        serialize_open_tag(node, out)
        if not node.children or #node.children == 0 then
          out[#out + 1] = "/>"
          stack[#stack] = nil
        else
          out[#out + 1] = ">"
          top[2] = 1
        end
      end
    elseif idx > #node.children then
      out[#out + 1] = "</"
      out[#out + 1] = node.tag
      out[#out + 1] = ">"
      stack[#stack] = nil
    else
      top[2] = idx + 1
      stack[#stack + 1] = { node.children[idx], 0 }
    end
  end
end

function M.serialize(root, opts)
  opts = opts or {}
  local out = {}
  if opts.declaration ~= false then
    out[#out + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  end
  serialize_node(root, out)
  return table.concat(out)
end

return M
