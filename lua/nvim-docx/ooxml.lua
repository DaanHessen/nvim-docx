-- Pure-Lua OOXML (WordprocessingML .docx) reader and writer for nvim-docx.
-- Supports styles preservation, media extraction/embedding, lists, tables,
-- hyperlinks, footnotes, endnotes, code blocks, blockquotes, track changes, and metadata.

local xml = require("nvim-docx.xml")
local image = require("nvim-docx.image")
local markdown = require("nvim-docx.markdown")

local M = {}

local NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
local NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
local NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
local NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
local NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
local NS_V = "urn:schemas-microsoft-com:vml"
local NS_O = "urn:schemas-microsoft-com:office:office"
local NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
local NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
local NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

local REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
local REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
local REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
local REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
local REL_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
local REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
local REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
local REL_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
local REL_WEB_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
local REL_FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
local REL_CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
local REL_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

local MONOSPACE_FONTS = {
  ["consolas"] = true,
  ["courier new"] = true,
  ["courier"] = true,
  ["monospace"] = true,
  ["lucida console"] = true,
  ["menlo"] = true,
  ["monaco"] = true,
  ["fira code"] = true,
  ["jetbrains mono"] = true,
  ["dejavu sans mono"] = true,
  ["source code pro"] = true,
}

local function is_bool_on(node)
  if not node then
    return false
  end
  local val = xml.get_attr(node, "w:val") or xml.get_attr(node, "val")
  if val == nil then
    return true
  end
  val = val:lower()
  return val ~= "0" and val ~= "false" and val ~= "off" and val ~= "none"
end

local function normalize_zip_path(base_dir, target)
  if not target or target == "" then
    return ""
  end
  if target:sub(1, 1) == "/" then
    return target:sub(2)
  end
  local combined = (base_dir and base_dir ~= "") and (base_dir .. "/" .. target) or target
  local parts = {}
  for part in combined:gmatch("[^/]+") do
    if part == ".." then
      if #parts > 0 then
        table.remove(parts)
      end
    elseif part ~= "." and part ~= "" then
      parts[#parts + 1] = part
    end
  end
  return table.concat(parts, "/")
end

local function write_binary_file(path, data)
  local dir = vim.fn.fnamemodify(path, ":h")
  if dir and dir ~= "" and vim.fn.isdirectory(dir) == 0 then
    vim.fn.mkdir(dir, "p")
  end
  local f = io.open(path, "wb")
  if not f then
    return false
  end
  f:write(data)
  f:close()
  return true
end

local function read_binary_file(path)
  local f = io.open(path, "rb")
  if not f then
    return nil
  end
  local data = f:read("*a")
  f:close()
  return data
end

--------------------------------------------------------------------------------
-- DOCX -> AST Reader
--------------------------------------------------------------------------------

local function parse_core_metadata(archive)
  local meta = {}
  local raw = archive.files["docProps/core.xml"]
  if not raw then
    return meta
  end
  local root = xml.parse(raw)
  if not root then
    return meta
  end
  local fields = {
    title = "title",
    creator = "author",
    subject = "subject",
    keywords = "keywords",
    description = "description",
  }
  for xml_name, meta_key in pairs(fields) do
    local child = xml.find_child(root, xml_name)
    if child then
      local txt = xml.get_direct_text(child):gsub("^%s+", ""):gsub("%s+$", "")
      if txt ~= "" then
        meta[meta_key] = txt
      end
    end
  end
  return meta
end

local function parse_relationships(archive, rels_path)
  local rels = {}
  local raw = archive.files[rels_path]
  if not raw then
    return rels
  end
  local root = xml.parse(raw)
  if not root then
    return rels
  end
  for _, rel in ipairs(xml.find_children(root, "Relationship")) do
    local id = xml.get_attr(rel, "Id")
    local rtype = xml.get_attr(rel, "Type") or ""
    local target = xml.get_attr(rel, "Target") or ""
    local mode = xml.get_attr(rel, "TargetMode") or ""
    if id then
      rels[id] = {
        id = id,
        type = rtype,
        target = target,
        target_mode = mode,
      }
    end
  end
  return rels
end

local function parse_styles(archive)
  local styles = {}
  local raw = archive.files["word/styles.xml"]
  if not raw then
    return styles
  end
  local root = xml.parse(raw)
  if not root then
    return styles
  end

  for _, snode in ipairs(xml.find_children(root, "style")) do
    local sid = xml.get_attr(snode, "w:styleId")
    local stype = xml.get_attr(snode, "w:type") or "paragraph"
    if sid then
      local name_node = xml.find_child(snode, "name")
      local sname = name_node and xml.get_attr(name_node, "w:val") or sid
      local based_node = xml.find_child(snode, "basedOn")
      local based_on = based_node and xml.get_attr(based_node, "w:val") or nil

      local outline_lvl = nil
      local num_id = nil
      local ilvl = nil
      local ppr = xml.find_child(snode, "pPr")
      if ppr then
        local ol = xml.find_child(ppr, "outlineLvl")
        if ol then
          outline_lvl = tonumber(xml.get_attr(ol, "w:val"))
        end
        local num_pr = xml.find_child(ppr, "numPr")
        if num_pr then
          local nid = xml.find_child(num_pr, "numId")
          local lvl = xml.find_child(num_pr, "ilvl")
          if nid then num_id = xml.get_attr(nid, "w:val") end
          if lvl then ilvl = tonumber(xml.get_attr(lvl, "w:val")) end
        end
      end

      local rpr = xml.find_child(snode, "rPr")
      local is_mono = false
      local is_bold = false
      local is_italic = false
      local is_strike = false
      if rpr then
        local rfonts = xml.find_child(rpr, "rFonts")
        if rfonts then
          local font = (xml.get_attr(rfonts, "w:ascii") or xml.get_attr(rfonts, "w:hAnsi") or ""):lower()
          if MONOSPACE_FONTS[font] then
            is_mono = true
          end
        end
        is_bold = is_bool_on(xml.find_child(rpr, "b"))
        is_italic = is_bool_on(xml.find_child(rpr, "i"))
        is_strike = is_bool_on(xml.find_child(rpr, "strike")) or is_bool_on(xml.find_child(rpr, "dstrike"))
      end

      styles[sid] = {
        id = sid,
        name = sname,
        type = stype,
        based_on = based_on,
        outline_lvl = outline_lvl,
        num_id = num_id,
        ilvl = ilvl,
        is_mono = is_mono,
        is_bold = is_bold,
        is_italic = is_italic,
        is_strike = is_strike,
      }
    end
  end

  return styles
end

local function resolve_heading_level(styles, style_id)
  if not style_id or style_id == "" then
    return nil
  end
  local visited = {}
  local cur = style_id
  while cur and not visited[cur] do
    visited[cur] = true
    local info = styles[cur]
    local candidates = { cur, info and info.name or nil }
    for _, cand in ipairs(candidates) do
      if cand then
        local lower = cand:lower()
        local h_num = lower:match("^heading%s*([1-9])$")
          or lower:match("^titre%s*([1-9])$")
          or lower:match("^kop%s*([1-9])$")
          or lower:match("berschrift%s*([1-9])$")
        if h_num then
          return math.min(tonumber(h_num), 6)
        end
        if lower == "title" then
          return 1
        end
        if lower == "subtitle" then
          return 2
        end
      end
    end
    if info then
      if info.outline_lvl and info.outline_lvl >= 0 and info.outline_lvl <= 8 then
        return math.min(info.outline_lvl + 1, 6)
      end
      cur = info.based_on
    else
      break
    end
  end
  return nil
end

local function is_code_para_style(styles, style_id)
  if not style_id or style_id == "" then
    return false
  end
  local visited = {}
  local cur = style_id
  while cur and not visited[cur] do
    visited[cur] = true
    local info = styles[cur]
    for _, cand in ipairs({ cur, info and info.name or nil }) do
      if cand then
        local l = cand:lower():gsub("%s+", "")
        if l == "sourcecode" or l == "code" or l == "codeblock" or l == "preformatted" or l == "htmlpreformatted" then
          return true
        end
      end
    end
    cur = info and info.based_on or nil
  end
  return false
end

local function is_blockquote_para_style(styles, style_id)
  if not style_id or style_id == "" then
    return false
  end
  local visited = {}
  local cur = style_id
  while cur and not visited[cur] do
    visited[cur] = true
    local info = styles[cur]
    for _, cand in ipairs({ cur, info and info.name or nil }) do
      if cand then
        local l = cand:lower():gsub("%s+", "")
        if l == "blocktext" or l == "quote" or l == "intensequote" or l == "blockquote" or l == "quotations" then
          return true
        end
      end
    end
    cur = info and info.based_on or nil
  end
  return false
end

local function is_code_char_style(styles, style_id)
  if not style_id or style_id == "" then
    return false
  end
  local visited = {}
  local cur = style_id
  while cur and not visited[cur] do
    visited[cur] = true
    local info = styles[cur]
    for _, cand in ipairs({ cur, info and info.name or nil }) do
      if cand then
        local l = cand:lower():gsub("%s+", "")
        if l == "verbatimchar" or l == "codechar" or l == "code" or l == "htmlcode" then
          return true
        end
      end
    end
    if info and info.is_mono then
      return true
    end
    cur = info and info.based_on or nil
  end
  return false
end

local function parse_numbering(archive)
  local numbering = {
    abstract = {},
    nums = {},
  }
  local raw = archive.files["word/numbering.xml"]
  if not raw then
    return numbering
  end
  local root = xml.parse(raw)
  if not root then
    return numbering
  end

  for _, anode in ipairs(xml.find_children(root, "abstractNum")) do
    local aid = xml.get_attr(anode, "w:abstractNumId")
    if aid then
      local levels = {}
      for _, lnode in ipairs(xml.find_children(anode, "lvl")) do
        local ilvl = tonumber(xml.get_attr(lnode, "w:ilvl")) or 0
        local fmt_node = xml.find_child(lnode, "numFmt")
        local start_node = xml.find_child(lnode, "start")
        local num_fmt = fmt_node and xml.get_attr(fmt_node, "w:val") or "bullet"
        local start_val = start_node and tonumber(xml.get_attr(start_node, "w:val")) or 1
        levels[ilvl] = {
          num_fmt = num_fmt,
          start = start_val,
        }
      end
      numbering.abstract[aid] = levels
    end
  end

  for _, nnode in ipairs(xml.find_children(root, "num")) do
    local nid = xml.get_attr(nnode, "w:numId")
    local aid_node = xml.find_child(nnode, "abstractNumId")
    local aid = aid_node and xml.get_attr(aid_node, "w:val") or nil
    if nid and aid then
      local overrides = {}
      for _, ov in ipairs(xml.find_children(nnode, "lvlOverride")) do
        local ilvl = tonumber(xml.get_attr(ov, "w:ilvl")) or 0
        local start_ov = xml.find_child(ov, "startOverride")
        local lvl_node = xml.find_child(ov, "lvl")
        local fmt_ov = nil
        if lvl_node then
          local fn = xml.find_child(lvl_node, "numFmt")
          if fn then fmt_ov = xml.get_attr(fn, "w:val") end
        end
        overrides[ilvl] = {
          start = start_ov and tonumber(xml.get_attr(start_ov, "w:val")) or nil,
          num_fmt = fmt_ov,
        }
      end
      numbering.nums[nid] = {
        abstract_id = aid,
        overrides = overrides,
      }
    end
  end

  return numbering
end

local function resolve_list_props(numbering, styles, num_id, ilvl, pstyle)
  -- Explicit <w:numId w:val="0"/> on the paragraph suppresses inherited list numbering
  if num_id and tostring(num_id) == "0" then
    return false, false, 0, 1, nil, false
  end
  if not num_id and pstyle and styles[pstyle] then
    num_id = styles[pstyle].num_id
    if not ilvl then
      ilvl = styles[pstyle].ilvl
    end
  end
  if not num_id or tostring(num_id) == "0" then
    -- Check if style name implies list
    if pstyle then
      local sl = pstyle:lower()
      if sl:find("listbullet") or sl == "listparagraph" then
        return true, false, ilvl or 0, 1, "style:" .. sl, false
      elseif sl:find("listnumber") then
        return true, true, ilvl or 0, 1, "style:" .. sl, false
      end
    end
    return false, false, 0, 1, nil, false
  end

  ilvl = tonumber(ilvl) or 0
  local ninfo = numbering.nums[tostring(num_id)]
  local num_fmt = "bullet"
  local start_val = 1
  local abs_key = "num:" .. tostring(num_id)
  local has_start_override = false
  if ninfo then
    abs_key = "abs:" .. tostring(ninfo.abstract_id)
    local abs = numbering.abstract[ninfo.abstract_id]
    if abs and abs[ilvl] then
      num_fmt = abs[ilvl].num_fmt or "bullet"
      start_val = abs[ilvl].start or 1
    end
    local ov = ninfo.overrides and ninfo.overrides[ilvl]
    if ov then
      if ov.num_fmt then num_fmt = ov.num_fmt end
      if ov.start then
        start_val = ov.start
        has_start_override = true
      end
    end
  end

  local ordered = (num_fmt ~= "bullet" and num_fmt ~= "none")
  return true, ordered, ilvl, start_val, abs_key, has_start_override
end

function M.docx_to_ast(archive, workspace, opts)
  opts = opts or {}
  local track_policy = opts.track_changes or "accept"

  -- Resolve main document path from _rels/.rels if present
  local root_rels = parse_relationships(archive, "_rels/.rels")
  local doc_path = "word/document.xml"
  for _, rel in pairs(root_rels) do
    if rel.type == REL_OFFICE_DOC and rel.target ~= "" then
      local candidate = normalize_zip_path("", rel.target)
      if archive.files[candidate] then
        doc_path = candidate
        break
      end
    end
  end

  local doc_xml_raw = archive.files[doc_path]
  if not doc_xml_raw then
    return nil, "Invalid DOCX archive: missing word/document.xml"
  end

  local doc_root, parse_err = xml.parse(doc_xml_raw)
  if not doc_root then
    return nil, "Invalid DOCX archive: malformed word/document.xml (" .. tostring(parse_err) .. ")"
  end

  local doc_dir = vim.fn.fnamemodify(doc_path, ":h")
  if doc_dir == "." then doc_dir = "" end
  local doc_rels_path = (doc_dir ~= "" and (doc_dir .. "/_rels/") or "_rels/")
    .. vim.fn.fnamemodify(doc_path, ":t") .. ".rels"

  local doc_rels = parse_relationships(archive, doc_rels_path)
  local footnote_rels = parse_relationships(archive, (doc_dir ~= "" and (doc_dir .. "/_rels/") or "_rels/") .. "footnotes.xml.rels")
  local endnote_rels = parse_relationships(archive, (doc_dir ~= "" and (doc_dir .. "/_rels/") or "_rels/") .. "endnotes.xml.rels")
  local styles = parse_styles(archive)
  local numbering = parse_numbering(archive)
  local metadata = parse_core_metadata(archive)

  -- Extract embedded media files to workspace/media/
  local rel_media_map = {}
  local function extract_rels_media(rels_table)
    for rid, rel in pairs(rels_table or {}) do
      if rel.type == REL_IMAGE and rel.target ~= "" then
        if rel.target_mode == "External" then
          rel_media_map[rid] = rel.target
        else
          local zip_entry_path = normalize_zip_path(doc_dir, rel.target)
          local fname = vim.fn.fnamemodify(zip_entry_path, ":t")
          local rel_out_path = "media/" .. fname
          rel_media_map[rid] = rel_out_path
          if workspace and workspace ~= "" and archive.files[zip_entry_path] then
            write_binary_file(workspace .. "/" .. rel_out_path, archive.files[zip_entry_path])
          end
        end
      end
    end
  end
  extract_rels_media(doc_rels)
  extract_rels_media(footnote_rels)
  extract_rels_media(endnote_rels)

  -- Also extract any other files in word/media/* into workspace/media/*
  if workspace and workspace ~= "" then
    for name, content in pairs(archive.files) do
      local media_sub = name:match("^word/(media/.+)$") or name:match("^(media/.+)$")
      if media_sub then
        write_binary_file(workspace .. "/" .. media_sub, content)
      end
    end
  end

  local ast = {
    metadata = metadata,
    blocks = {},
    footnotes = {},
  }

  -- Pre-parse footnotes.xml, endnotes.xml, and comments.xml
  local raw_footnotes = {}
  local raw_endnotes = {}
  local raw_comments = {}

  local function extract_omath_text(mnode)
    if not mnode then
      return ""
    end
    local lname = mnode.local_name
    if lname == "t" then
      return xml.get_direct_text(mnode)
    elseif lname == "f" then
      local num = extract_omath_text(xml.find_child(mnode, "num"))
      local den = extract_omath_text(xml.find_child(mnode, "den"))
      return "\\frac{" .. num .. "}{" .. den .. "}"
    elseif lname == "sSup" then
      local base = extract_omath_text(xml.find_child(mnode, "e"))
      local sup = extract_omath_text(xml.find_child(mnode, "sup"))
      return base .. "^{" .. sup .. "}"
    elseif lname == "sSub" then
      local base = extract_omath_text(xml.find_child(mnode, "e"))
      local sub = extract_omath_text(xml.find_child(mnode, "sub"))
      return base .. "_{" .. sub .. "}"
    elseif lname == "sSubSup" then
      local base = extract_omath_text(xml.find_child(mnode, "e"))
      local sub = extract_omath_text(xml.find_child(mnode, "sub"))
      local sup = extract_omath_text(xml.find_child(mnode, "sup"))
      return base .. "_{" .. sub .. "}^{" .. sup .. "}"
    elseif lname == "rad" then
      local deg = extract_omath_text(xml.find_child(mnode, "deg")):gsub("^%s+", ""):gsub("%s+$", "")
      local base = extract_omath_text(xml.find_child(mnode, "e"))
      if deg ~= "" then
        return "\\sqrt[" .. deg .. "]{" .. base .. "}"
      end
      return "\\sqrt{" .. base .. "}"
    elseif lname == "d" then
      local dpr = xml.find_child(mnode, "dPr")
      local beg_chr = "("
      local end_chr = ")"
      local sep_chr = ","
      if dpr then
        local bc = xml.find_child(dpr, "begChr")
        local ec = xml.find_child(dpr, "endChr")
        local sc = xml.find_child(dpr, "sepChr")
        if bc and xml.get_attr(bc, "m:val") then beg_chr = xml.get_attr(bc, "m:val") end
        if ec and xml.get_attr(ec, "m:val") then end_chr = xml.get_attr(ec, "m:val") end
        if sc and xml.get_attr(sc, "m:val") then sep_chr = xml.get_attr(sc, "m:val") end
      end
      local parts = {}
      for _, echild in ipairs(xml.find_children(mnode, "e")) do
        parts[#parts + 1] = extract_omath_text(echild)
      end
      return beg_chr .. table.concat(parts, sep_chr) .. end_chr
    end

    local out = {}
    for _, c in ipairs(mnode.children or {}) do
      if c.tag then
        out[#out + 1] = extract_omath_text(c)
      end
    end
    return table.concat(out)
  end

  local function extract_run_drawing(drawing_node, active_rels)
    -- Check for horizontal rule VML rect
    for _, rect in ipairs(xml.find_all(drawing_node, "rect")) do
      if xml.get_attr(rect, "o:hr") == "t" or xml.get_attr(rect, "hr") == "t" then
        return { is_hr = true }
      end
    end

    local r_embed = nil
    for _, blip in ipairs(xml.find_all(drawing_node, "blip")) do
      r_embed = xml.get_attr(blip, "r:embed") or xml.get_attr(blip, "r:link")
      if r_embed then break end
    end
    if not r_embed then
      for _, imgdata in ipairs(xml.find_all(drawing_node, "imagedata")) do
        r_embed = xml.get_attr(imgdata, "r:id") or xml.get_attr(imgdata, "o:relid")
        if r_embed then break end
      end
    end

    if not r_embed then
      return nil
    end

    local src = rel_media_map[r_embed]
    if not src and active_rels and active_rels[r_embed] then
      src = active_rels[r_embed].target
    end
    if not src and doc_rels[r_embed] then
      src = doc_rels[r_embed].target
    end
    if not src then
      return nil
    end

    local alt = ""
    for _, doc_pr in ipairs(xml.find_all(drawing_node, "docPr")) do
      local descr = xml.get_attr(doc_pr, "descr") or xml.get_attr(doc_pr, "title")
      if descr and descr ~= "" then
        alt = descr
        break
      end
    end
    if alt == "" then
      for _, cnv in ipairs(xml.find_all(drawing_node, "cNvPr")) do
        local descr = xml.get_attr(cnv, "descr")
        if descr and descr ~= "" then
          alt = descr
          break
        end
      end
    end

    local width_str, height_str = nil, nil
    local extent = xml.find_all(drawing_node, "extent")[1]
    if extent then
      local cx = xml.get_attr(extent, "cx")
      local cy = xml.get_attr(extent, "cy")
      width_str = image.emu_to_inches_str(cx)
      height_str = image.emu_to_inches_str(cy)
    else
      -- Check VML shape style="...width:126.25pt;height:187.35pt..."
      for _, el in ipairs(xml.find_all(drawing_node)) do
        local st = xml.get_attr(el, "style")
        if st then
          local w_val = st:match("width%s*:%s*([^;%s]+)")
          local h_val = st:match("height%s*:%s*([^;%s]+)")
          if w_val or h_val then
            if w_val then width_str = image.emu_to_inches_str(image.parse_length_to_emu(w_val)) end
            if h_val then height_str = image.emu_to_inches_str(image.parse_length_to_emu(h_val)) end
            break
          end
        end
      end
    end

    return {
      type = "image",
      src = src,
      alt = alt,
      width = width_str,
      height = height_str,
    }
  end

  local fn_order_counter = 0
  local fn_id_remap = {}
  local referenced_footnotes = {}

  local function register_footnote_ref(kind, orig_id)
    local key = kind .. ":" .. tostring(orig_id)
    if not fn_id_remap[key] then
      fn_order_counter = fn_order_counter + 1
      local disp_id = tostring(fn_order_counter)
      fn_id_remap[key] = disp_id
      referenced_footnotes[#referenced_footnotes + 1] = {
        key = key,
        kind = kind,
        orig_id = tostring(orig_id),
        disp_id = disp_id,
      }
    end
    return fn_id_remap[key]
  end

  local function parse_hyperlink_instr(instr)
    if not instr or instr == "" then
      return nil
    end
    local anchor = instr:match('HYPERLINK%s+\\l%s+"([^"]+)"')
    if anchor and anchor ~= "" then
      return "#" .. anchor
    end
    local url = instr:match('HYPERLINK%s+"([^"]+)"')
    if url and url ~= "" then
      local sub_anchor = instr:match('\\l%s+"([^"]+)"')
      if sub_anchor and sub_anchor ~= "" then
        return url .. "#" .. sub_anchor
      end
      return url
    end
    return nil
  end

  local function extract_inlines_from_nodes(nodes, in_code_para, inherited_fmt, active_rels)
    local inlines = {}
    local found_hr = false
    local found_checkbox = nil
    local only_omath_para = false
    local omath_para_count = 0
    local other_element_count = 0
    inherited_fmt = inherited_fmt or {}
    active_rels = active_rels or doc_rels

    local field_stack = {}

    local function push_inline(node)
      if #field_stack > 0 then
        local top = field_stack[#field_stack]
        if top.state == "result" then
          top.inlines[#top.inlines + 1] = node
        end
      else
        inlines[#inlines + 1] = node
      end
    end

    local function close_field()
      local top = table.remove(field_stack)
      if not top then
        return
      end
      local hurl = parse_hyperlink_instr(top.instr)
      local merged_sub = markdown.merge_inlines(top.inlines)
      if hurl then
        if #merged_sub == 0 then
          merged_sub = { { type = "text", text = hurl } }
        end
        push_inline({
          type = "link",
          url = hurl,
          inlines = merged_sub,
        })
      else
        for _, item in ipairs(merged_sub) do
          push_inline(item)
        end
      end
    end

    local function walk_inline_children(child_list, cur_fmt, in_del)
      for _, child in ipairs(child_list or {}) do
        local lname = child.local_name

        if lname == "r" then
          other_element_count = other_element_count + 1
          local rfmt = {
            bold = cur_fmt.bold or false,
            italic = cur_fmt.italic or false,
            strike = cur_fmt.strike or in_del or false,
            underline = cur_fmt.underline or false,
            code = cur_fmt.code or false,
            sup = cur_fmt.sup or false,
            sub = cur_fmt.sub or false,
          }

          if not in_code_para then
            local rpr = xml.find_child(child, "rPr")
            if rpr then
              local rstyle_node = xml.find_child(rpr, "rStyle")
              local rstyle = rstyle_node and xml.get_attr(rstyle_node, "w:val") or nil
              if rstyle then
                if is_code_char_style(styles, rstyle) then
                  rfmt.code = true
                end
                local sinfo = styles[rstyle]
                if sinfo then
                  if sinfo.is_bold or rstyle == "Strong" then rfmt.bold = true end
                  if sinfo.is_italic or rstyle == "Emphasis" then rfmt.italic = true end
                  if sinfo.is_strike then rfmt.strike = true end
                end
              end

              local b_node = xml.find_child(rpr, "b")
              if b_node then rfmt.bold = is_bool_on(b_node) end

              local i_node = xml.find_child(rpr, "i")
              if i_node then rfmt.italic = is_bool_on(i_node) end

              local s_node = xml.find_child(rpr, "strike") or xml.find_child(rpr, "dstrike")
              if s_node then rfmt.strike = is_bool_on(s_node) end

              local va_node = xml.find_child(rpr, "vertAlign")
              if va_node then
                local vval = (xml.get_attr(va_node, "w:val") or ""):lower()
                if vval == "superscript" then
                  rfmt.sup = true
                elseif vval == "subscript" then
                  rfmt.sub = true
                end
              end

              local rfonts = xml.find_child(rpr, "rFonts")
              if rfonts then
                local f = (xml.get_attr(rfonts, "w:ascii") or xml.get_attr(rfonts, "w:hAnsi") or ""):lower()
                if MONOSPACE_FONTS[f] then
                  rfmt.code = true
                end
              end
            end
          end

          for _, rchild in ipairs(child.children or {}) do
            local rlname = rchild.local_name
            if rlname == "fldChar" then
              local ftype = (xml.get_attr(rchild, "w:fldCharType") or ""):lower()
              if ftype == "begin" then
                field_stack[#field_stack + 1] = {
                  state = "instr",
                  instr = "",
                  inlines = {},
                }
              elseif ftype == "separate" then
                if #field_stack > 0 then
                  field_stack[#field_stack].state = "result"
                end
              elseif ftype == "end" then
                close_field()
              end
            elseif rlname == "instrText" then
              if #field_stack > 0 then
                field_stack[#field_stack].instr = field_stack[#field_stack].instr .. xml.get_direct_text(rchild)
              end
            elseif rlname == "t" or (rlname == "delText" and in_del) then
              local txt = xml.get_direct_text(rchild)
              if txt ~= "" then
                local node = vim.deepcopy(rfmt)
                node.type = "text"
                node.text = txt
                push_inline(node)
              end
            elseif rlname == "tab" then
              if in_code_para then
                push_inline({ type = "text", text = "\t" })
              else
                push_inline({ type = "tab" })
              end
            elseif rlname == "br" or rlname == "cr" then
              local br_type = xml.get_attr(rchild, "w:type")
              if br_type == "page" then
                found_hr = true
              elseif in_code_para then
                push_inline({ type = "text", text = "\n" })
              else
                push_inline({ type = "linebreak" })
              end
            elseif rlname == "noBreakHyphen" then
              local node = vim.deepcopy(rfmt)
              node.type = "text"
              node.text = "-"
              push_inline(node)
            elseif rlname == "drawing" or rlname == "pict" then
              local d = extract_run_drawing(rchild, active_rels)
              if d then
                if d.is_hr then
                  found_hr = true
                else
                  push_inline(d)
                end
              end
            elseif rlname == "footnoteReference" then
              local fn_id = xml.get_attr(rchild, "w:id")
              if fn_id and tonumber(fn_id) and tonumber(fn_id) > 0 then
                local disp = register_footnote_ref("fn", fn_id)
                push_inline({ type = "footnote_ref", id = disp })
              end
            elseif rlname == "endnoteReference" then
              local en_id = xml.get_attr(rchild, "w:id")
              if en_id and tonumber(en_id) and tonumber(en_id) > 0 then
                local disp = register_footnote_ref("en", en_id)
                push_inline({ type = "footnote_ref", id = disp })
              end
            elseif rlname == "commentReference" then
              if track_policy == "all" then
                local cm_id = xml.get_attr(rchild, "w:id")
                if cm_id and raw_comments[tostring(cm_id)] then
                  local disp = register_footnote_ref("cm", cm_id)
                  push_inline({ type = "footnote_ref", id = disp })
                end
              end
            end
          end

        elseif lname == "commentReference" then
          if track_policy == "all" then
            local cm_id = xml.get_attr(child, "w:id")
            if cm_id and raw_comments[tostring(cm_id)] then
              local disp = register_footnote_ref("cm", cm_id)
              push_inline({ type = "footnote_ref", id = disp })
            end
          end

        elseif lname == "oMathPara" or lname == "oMath" then
          if lname == "oMathPara" then
            omath_para_count = omath_para_count + 1
          else
            other_element_count = other_element_count + 1
          end
          local mtxt = extract_omath_text(child):gsub("^%s+", ""):gsub("%s+$", "")
          if mtxt ~= "" then
            push_inline({
              type = "math",
              text = mtxt,
            })
          end

        elseif lname == "hyperlink" then
          other_element_count = other_element_count + 1
          local rid = xml.get_attr(child, "r:id")
          local anchor = xml.get_attr(child, "w:anchor")
          local url = ""
          if rid and active_rels and active_rels[rid] then
            url = active_rels[rid].target or ""
          elseif rid and doc_rels[rid] then
            url = doc_rels[rid].target or ""
          elseif anchor and anchor ~= "" then
            url = "#" .. anchor
          end
          local link_children = extract_inlines_from_nodes(child.children or {}, false, cur_fmt, active_rels)
          if url ~= "" then
            push_inline({
              type = "link",
              url = url,
              inlines = link_children,
            })
          else
            for _, lc in ipairs(link_children) do
              push_inline(lc)
            end
          end

        elseif lname == "fldSimple" then
          other_element_count = other_element_count + 1
          local instr = xml.get_attr(child, "w:instr") or ""
          local hurl = parse_hyperlink_instr(instr)
          local sub_inlines = extract_inlines_from_nodes(child.children or {}, false, cur_fmt, active_rels)
          if hurl then
            push_inline({
              type = "link",
              url = hurl,
              inlines = sub_inlines,
            })
          else
            for _, si in ipairs(sub_inlines) do
              push_inline(si)
            end
          end

        elseif lname == "ins" or lname == "moveTo" then
          if track_policy ~= "reject" then
            walk_inline_children(child.children, cur_fmt, false)
          end

        elseif lname == "del" or lname == "moveFrom" then
          if track_policy == "reject" then
            walk_inline_children(child.children, cur_fmt, true)
          elseif track_policy == "all" then
            local del_fmt = vim.deepcopy(cur_fmt)
            del_fmt.strike = true
            walk_inline_children(child.children, del_fmt, true)
          end

        elseif lname == "sdt" then
          local sdt_pr = xml.find_child(child, "sdtPr")
          if sdt_pr then
            local cb = xml.find_all(sdt_pr, "checkbox")[1]
            if cb then
              local checked_node = xml.find_all(cb, "checked")[1]
              if checked_node then
                found_checkbox = is_bool_on(checked_node)
              else
                found_checkbox = false
              end
            end
          end
          local sdt_content = xml.find_child(child, "sdtContent")
          if sdt_content then
            walk_inline_children(sdt_content.children, cur_fmt, in_del)
          end

        elseif lname == "smartTag" or lname == "customXml" then
          walk_inline_children(child.children, cur_fmt, in_del)
        end
      end
    end

    walk_inline_children(nodes, inherited_fmt, false)
    while #field_stack > 0 do
      close_field()
    end
    if omath_para_count == 1 and other_element_count == 0 then
      only_omath_para = true
    end
    return markdown.merge_inlines(inlines), found_hr, found_checkbox, only_omath_para
  end

  local function trim_inlines(inlines, trim_leading_spaces)
    if not inlines then
      return {}
    end
    while #inlines > 0 do
      local first = inlines[1]
      if first.type == "tab" then
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
      if last.type == "tab" then
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
    return inlines
  end

  local function parse_notes_file(path, target_table, tag_name, note_rels)
    local raw = archive.files[path]
    if not raw then return end
    local root = xml.parse(raw)
    if not root then return end
    for _, fnode in ipairs(xml.find_children(root, tag_name)) do
      local fid = xml.get_attr(fnode, "w:id")
      local ftype = xml.get_attr(fnode, "w:type")
      local is_valid_id = fid and tonumber(fid) ~= nil and (tag_name == "comment" and tonumber(fid) >= 0 or tonumber(fid) > 0)
      if is_valid_id and ftype ~= "separator" and ftype ~= "continuationSeparator" then
        local note_inlines = {}
        local paras = xml.find_children(fnode, "p")
        for pi, p in ipairs(paras) do
          local pinl = trim_inlines(extract_inlines_from_nodes(p.children, false, nil, note_rels), true)
          for _, item in ipairs(pinl) do
            note_inlines[#note_inlines + 1] = item
          end
          if pi < #paras and #pinl > 0 then
            note_inlines[#note_inlines + 1] = { type = "text", text = " " }
          end
        end
        local merged_notes = trim_inlines(markdown.merge_inlines(note_inlines), true)
        if tag_name == "comment" then
          local author = xml.get_attr(fnode, "w:author")
          if author and author ~= "" then
            table.insert(merged_notes, 1, { type = "text", text = "[" .. author .. "]: ", bold = true })
            merged_notes = markdown.merge_inlines(merged_notes)
          end
        end
        target_table[tostring(fid)] = merged_notes
      end
    end
  end

  parse_notes_file("word/comments.xml", raw_comments, "comment", doc_rels)
  parse_notes_file("word/footnotes.xml", raw_footnotes, "footnote", footnote_rels)
  parse_notes_file("word/endnotes.xml", raw_endnotes, "endnote", endnote_rels)

  local ordered_counters = {}
  local seen_override_nums = {}

  local function process_paragraph(pnode)
    local ppr = xml.find_child(pnode, "pPr")
    local pstyle = nil
    local direct_outline_lvl = nil
    local num_id = nil
    local ilvl = nil
    local has_bottom_border = false

    if ppr then
      local ps = xml.find_child(ppr, "pStyle")
      if ps then pstyle = xml.get_attr(ps, "w:val") end

      local ol = xml.find_child(ppr, "outlineLvl")
      if ol then direct_outline_lvl = tonumber(xml.get_attr(ol, "w:val")) end

      local num_pr = xml.find_child(ppr, "numPr")
      if num_pr then
        local nid = xml.find_child(num_pr, "numId")
        local lvl = xml.find_child(num_pr, "ilvl")
        if nid then num_id = xml.get_attr(nid, "w:val") end
        if lvl then ilvl = tonumber(xml.get_attr(lvl, "w:val")) end
      end

      local pbdr = xml.find_child(ppr, "pBdr")
      if pbdr and xml.find_child(pbdr, "bottom") then
        has_bottom_border = true
      end
    end

    local is_code = is_code_para_style(styles, pstyle)
    local inlines, found_hr, found_checkbox, only_omath_para = extract_inlines_from_nodes(pnode.children, is_code)

    if is_code then
      local raw_parts = {}
      for _, inl in ipairs(inlines) do
        if inl.type == "text" then
          raw_parts[#raw_parts + 1] = inl.text
        elseif inl.type == "linebreak" then
          raw_parts[#raw_parts + 1] = "\n"
        elseif inl.type == "tab" then
          raw_parts[#raw_parts + 1] = "\t"
        end
      end
      local line_code = table.concat(raw_parts)
      local prev_block = ast.blocks[#ast.blocks]
      if prev_block and prev_block.type == "code_block" then
        prev_block.text = prev_block.text .. "\n" .. line_code
      else
        ast.blocks[#ast.blocks + 1] = {
          type = "code_block",
          lang = "",
          text = line_code,
        }
      end
      return
    end

    inlines = trim_inlines(inlines, false)

    if found_hr or (has_bottom_border and #inlines == 0) then
      ast.blocks[#ast.blocks + 1] = { type = "hr" }
      if #inlines == 0 then
        return
      end
    end

    if #inlines == 0 then
      return
    end

    if only_omath_para and #inlines == 1 and inlines[1].type == "math" then
      ast.blocks[#ast.blocks + 1] = {
        type = "math_block",
        text = inlines[1].text,
      }
      return
    end

    local h_lvl = resolve_heading_level(styles, pstyle)
    if not h_lvl and direct_outline_lvl and direct_outline_lvl >= 0 and direct_outline_lvl <= 8 then
      h_lvl = math.min(direct_outline_lvl + 1, 6)
    end

    if h_lvl then
      ordered_counters = {}
      seen_override_nums = {}
      ast.blocks[#ast.blocks + 1] = {
        type = "heading",
        level = h_lvl,
        inlines = trim_inlines(inlines, true),
      }
      return
    end

    local is_list, is_ordered, list_lvl, start_val, abs_key, has_start_override =
      resolve_list_props(numbering, styles, num_id, ilvl, pstyle)
    if is_list then
      inlines = trim_inlines(inlines, true)
      -- Check leading Unicode checkbox characters ☐ (E2 98 90), ☑ (E2 98 91), ☒ (E2 98 92)
      local first_inl = inlines[1]
      if first_inl and first_inl.type == "text" then
        if first_inl.text:sub(1, 3) == "\226\152\144" then
          found_checkbox = false
          first_inl.text = first_inl.text:sub(4):gsub("^%s+", "")
        elseif first_inl.text:sub(1, 3) == "\226\152\145" or first_inl.text:sub(1, 3) == "\226\152\146" then
          found_checkbox = true
          first_inl.text = first_inl.text:sub(4):gsub("^%s+", "")
        end
        if first_inl.text == "" then
          table.remove(inlines, 1)
        end
      end

      -- Clear counters for any deeper nesting levels so sublists restart at 1
      local list_prefix = tostring(abs_key or num_id or pstyle or "default") .. ":"
      for k, _ in pairs(ordered_counters) do
        local klvl = tonumber(k:match(":(%d+)$"))
        if klvl and klvl > list_lvl then
          ordered_counters[k] = nil
        end
      end

      local counter_key = list_prefix .. tostring(list_lvl)
      local override_key = tostring(num_id or "") .. ":" .. tostring(list_lvl)
      local item_num = nil
      if is_ordered then
        if has_start_override and not seen_override_nums[override_key] then
          seen_override_nums[override_key] = true
          ordered_counters[counter_key] = start_val or 1
        elseif not ordered_counters[counter_key] then
          ordered_counters[counter_key] = start_val or 1
        else
          ordered_counters[counter_key] = ordered_counters[counter_key] + 1
        end
        item_num = ordered_counters[counter_key]
      end

      ast.blocks[#ast.blocks + 1] = {
        type = "list_item",
        ordered = is_ordered,
        level = list_lvl,
        number = item_num,
        checked = found_checkbox,
        inlines = inlines,
      }
      return
    end

    ordered_counters = {}
    seen_override_nums = {}

    if is_blockquote_para_style(styles, pstyle) then
      ast.blocks[#ast.blocks + 1] = {
        type = "blockquote",
        level = 1,
        inlines = inlines,
      }
      return
    end

    ast.blocks[#ast.blocks + 1] = {
      type = "paragraph",
      inlines = inlines,
    }
  end

  local function process_table(tbl_node)
    ordered_counters = {}
    seen_override_nums = {}
    local all_rows = {}
    local alignments = {}

    local function collect_tr(container)
      for _, child in ipairs(container.children or {}) do
        if child.local_name == "tr" then
          local row_cells = {}
          local col_idx = 0
          for _, tc in ipairs(child.children or {}) do
            if tc.local_name == "tc" then
              col_idx = col_idx + 1
              local cell_inlines = {}
              local paras = xml.find_children(tc, "p")
              for pi, p in ipairs(paras) do
                local ppr = xml.find_child(p, "pPr")
                if ppr and not alignments[col_idx] then
                  local jc = xml.find_child(ppr, "jc")
                  if jc then
                    local jval = (xml.get_attr(jc, "w:val") or ""):lower()
                    if jval == "center" or jval == "right" or jval == "left" then
                      alignments[col_idx] = jval
                    end
                  end
                end
                local pinl = trim_inlines(extract_inlines_from_nodes(p.children, false), true)
                if #pinl > 0 then
                  if #cell_inlines > 0 then
                    cell_inlines[#cell_inlines + 1] = { type = "linebreak" }
                  end
                  for _, item in ipairs(pinl) do
                    cell_inlines[#cell_inlines + 1] = item
                  end
                end
              end
              row_cells[#row_cells + 1] = trim_inlines(markdown.merge_inlines(cell_inlines), true)

              -- Handle horizontal span <w:gridSpan w:val="N"/>
              local tc_pr = xml.find_child(tc, "tcPr")
              if tc_pr then
                local gspan = xml.find_child(tc_pr, "gridSpan")
                local span_n = gspan and tonumber(xml.get_attr(gspan, "w:val")) or 1
                if span_n and span_n > 1 then
                  for _ = 2, span_n do
                    col_idx = col_idx + 1
                    row_cells[#row_cells + 1] = {}
                  end
                end
              end
            end
          end
          if #row_cells > 0 then
            all_rows[#all_rows + 1] = row_cells
          end
        elseif child.local_name == "sdt" then
          local sdt_c = xml.find_child(child, "sdtContent")
          if sdt_c then collect_tr(sdt_c) end
        end
      end
    end

    collect_tr(tbl_node)
    if #all_rows == 0 then
      return
    end

    local num_cols = 0
    for _, r in ipairs(all_rows) do
      if #r > num_cols then num_cols = #r end
    end
    for c = 1, num_cols do
      if not alignments[c] then alignments[c] = "left" end
    end

    local headers = all_rows[1]
    local data_rows = {}
    for r = 2, #all_rows do
      data_rows[#data_rows + 1] = all_rows[r]
    end

    ast.blocks[#ast.blocks + 1] = {
      type = "table",
      alignments = alignments,
      headers = headers,
      rows = data_rows,
    }
  end

  local function walk_body(container)
    if not container or not container.children then
      return
    end
    for _, child in ipairs(container.children) do
      local lname = child.local_name
      if lname == "p" then
        process_paragraph(child)
      elseif lname == "tbl" then
        process_table(child)
      elseif lname == "oMathPara" or lname == "oMath" then
        local mtxt = extract_omath_text(child):gsub("^%s+", ""):gsub("%s+$", "")
        if mtxt ~= "" then
          ast.blocks[#ast.blocks + 1] = {
            type = "math_block",
            text = mtxt,
          }
        end
      elseif lname == "sdt" then
        local sdt_content = xml.find_child(child, "sdtContent")
        if sdt_content then
          walk_body(sdt_content)
        end
      elseif lname == "customXml" or lname == "smartTag" then
        walk_body(child)
      elseif lname == "ins" or lname == "moveTo" then
        if track_policy ~= "reject" then
          walk_body(child)
        end
      elseif lname == "del" or lname == "moveFrom" then
        if track_policy == "reject" or track_policy == "all" then
          walk_body(child)
        end
      end
    end
  end

  local body = xml.find_child(doc_root, "body") or doc_root
  walk_body(body)

  -- Populate referenced footnotes/endnotes/comments in appearance order
  for _, ref in ipairs(referenced_footnotes) do
    local note_inlines = nil
    if ref.kind == "fn" then
      note_inlines = raw_footnotes[ref.orig_id]
    elseif ref.kind == "en" then
      note_inlines = raw_endnotes[ref.orig_id]
    elseif ref.kind == "cm" then
      note_inlines = raw_comments[ref.orig_id]
    end
    if note_inlines and #note_inlines > 0 then
      ast.footnotes[#ast.footnotes + 1] = {
        id = ref.disp_id,
        inlines = note_inlines,
      }
    end
  end

  return ast
end

--------------------------------------------------------------------------------
-- AST -> DOCX Writer
--------------------------------------------------------------------------------

local function default_styles_xml()
  return table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
    '<w:styles xmlns:w="', NS_W, '" xmlns:r="', NS_R, '">',
    '<w:docDefaults>',
      '<w:rPrDefault><w:rPr>',
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>',
        '<w:sz w:val="22"/><w:szCs w:val="22"/>',
        '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>',
      '</w:rPr></w:rPrDefault>',
      '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>',
    '</w:docDefaults>',
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">',
      '<w:name w:val="Normal"/><w:qFormat/>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="FirstParagraph">',
      '<w:name w:val="First Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Compact">',
      '<w:name w:val="Compact"/><w:basedOn w:val="Normal"/>',
      '<w:pPr><w:spacing w:before="0" w:after="36"/></w:pPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading1">',
      '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>',
      '<w:rPr><w:b/><w:bCs/><w:color w:val="2F5496"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading2">',
      '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr>',
      '<w:rPr><w:b/><w:bCs/><w:color w:val="2F5496"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading3">',
      '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr>',
      '<w:rPr><w:b/><w:bCs/><w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading4">',
      '<w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="40"/><w:outlineLvl w:val="3"/></w:pPr>',
      '<w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:color w:val="2F5496"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading5">',
      '<w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="40"/><w:outlineLvl w:val="4"/></w:pPr>',
      '<w:rPr><w:color w:val="2F5496"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading6">',
      '<w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="40"/><w:outlineLvl w:val="5"/></w:pPr>',
      '<w:rPr><w:i/><w:iCs/><w:color w:val="1F3763"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="BlockText">',
      '<w:name w:val="Block Text"/><w:basedOn w:val="Normal"/>',
      '<w:pPr><w:ind w:left="720" w:right="720"/></w:pPr>',
      '<w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="SourceCode">',
      '<w:name w:val="Source Code"/><w:basedOn w:val="Normal"/>',
      '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:shd w:val="clear" w:color="auto" w:fill="F6F8FA"/></w:pPr>',
      '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="character" w:styleId="VerbatimChar">',
      '<w:name w:val="Verbatim Char"/>',
      '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="character" w:styleId="Hyperlink">',
      '<w:name w:val="Hyperlink"/>',
      '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="character" w:styleId="FootnoteReference">',
      '<w:name w:val="footnote reference"/>',
      '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="FootnoteText">',
      '<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>',
      '<w:pPr><w:spacing w:after="60"/></w:pPr>',
      '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="table" w:default="1" w:styleId="TableGrid">',
      '<w:name w:val="Table Grid"/>',
      '<w:tblPr>',
        '<w:tblBorders>',
          '<w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>',
          '<w:left w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>',
          '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>',
          '<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>',
          '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>',
          '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>',
        '</w:tblBorders>',
      '</w:tblPr>',
    '</w:style>',
    '</w:styles>',
  })
end

local function ensure_required_styles(styles_xml_str)
  if not styles_xml_str or styles_xml_str == "" then
    return default_styles_xml(), {}
  end
  local root = xml.parse(styles_xml_str)
  if not root then
    return default_styles_xml(), {}
  end

  local existing = {}
  local heading_ids = {}
  for _, snode in ipairs(xml.find_children(root, "style")) do
    local sid = xml.get_attr(snode, "w:styleId")
    if sid then
      existing[sid] = true
      local name_node = xml.find_child(snode, "name")
      local sname = (name_node and xml.get_attr(name_node, "w:val") or sid):lower()
      local h_lvl = sname:match("^heading%s*([1-6])$") or sid:lower():match("^heading([1-6])$")
      if h_lvl then
        heading_ids[tonumber(h_lvl)] = sid
      end
    end
  end

  local def_root = xml.parse(default_styles_xml())
  local modified = false
  if def_root then
    for _, def_style in ipairs(xml.find_children(def_root, "style")) do
      local sid = xml.get_attr(def_style, "w:styleId")
      local h_lvl = sid and tonumber(sid:match("^Heading([1-6])$")) or nil
      if h_lvl then
        if not heading_ids[h_lvl] and not existing[sid] then
          root.children[#root.children + 1] = def_style
          heading_ids[h_lvl] = sid
          modified = true
        end
      elseif sid and not existing[sid] then
        root.children[#root.children + 1] = def_style
        modified = true
      end
    end
  end

  if modified then
    return xml.serialize(root), heading_ids
  end
  return styles_xml_str, heading_ids
end

local function build_numbering_xml(list_defs, existing_numbering_xml)
  local abs_children = {}
  local num_children = {}

  if existing_numbering_xml and existing_numbering_xml ~= "" then
    local ex_root = xml.parse(existing_numbering_xml)
    if ex_root then
      for _, c in ipairs(ex_root.children or {}) do
        if c.local_name == "abstractNum" then
          abs_children[#abs_children + 1] = c
        elseif c.local_name == "num" then
          num_children[#num_children + 1] = c
        end
      end
    end
  end

  local bullet_symbols = { "\226\128\162", "o", "\226\150\170", "\226\128\162", "o", "\226\150\170", "\226\128\162", "o", "\226\150\170" }
  local ordered_fmts = { "decimal", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman" }

  for _, def in ipairs(list_defs) do
    local lvls = {}
    for lvl = 0, 8 do
      local left = tostring(720 * (lvl + 1))
      local fmt_val = def.ordered and ordered_fmts[lvl + 1] or "bullet"
      local lvl_text = def.ordered and ("%" .. tostring(lvl + 1) .. ".") or bullet_symbols[lvl + 1]
      local start_val = (lvl == 0 and def.start) and tostring(def.start) or "1"

      local lvl_children = {
        xml.elem("w:start", { ["w:val"] = start_val }),
        xml.elem("w:numFmt", { ["w:val"] = fmt_val }),
        xml.elem("w:lvlText", { ["w:val"] = lvl_text }),
        xml.elem("w:lvlJc", { ["w:val"] = "left" }),
        xml.elem("w:pPr", nil, {
          xml.elem("w:ind", { ["w:left"] = left, ["w:hanging"] = "360" }),
        }),
      }
      if not def.ordered then
        lvl_children[#lvl_children + 1] = xml.elem("w:rPr", nil, {
          xml.elem("w:rFonts", { ["w:ascii"] = "Symbol", ["w:hAnsi"] = "Symbol", ["w:hint"] = "default" }),
        })
      end
      lvls[#lvls + 1] = xml.elem("w:lvl", { ["w:ilvl"] = tostring(lvl) }, lvl_children)
    end
    abs_children[#abs_children + 1] = xml.elem("w:abstractNum", { ["w:abstractNumId"] = tostring(def.abstract_id) }, lvls)
  end

  for _, def in ipairs(list_defs) do
    num_children[#num_children + 1] = xml.elem("w:num", { ["w:numId"] = tostring(def.num_id) }, {
      xml.elem("w:abstractNumId", { ["w:val"] = tostring(def.abstract_id) }),
    })
  end

  local children = {}
  for _, a in ipairs(abs_children) do
    children[#children + 1] = a
  end
  for _, n in ipairs(num_children) do
    children[#children + 1] = n
  end

  local root = xml.elem("w:numbering", {
    ["xmlns:w"] = NS_W,
    ["xmlns:r"] = NS_R,
  }, children)
  return xml.serialize(root)
end

local function build_drawing_element(r_id, pic_id, name, alt, cx, cy)
  cx = tostring(math.floor(tonumber(cx) or 2743200))
  cy = tostring(math.floor(tonumber(cy) or 1828800))
  pic_id = tostring(pic_id or 1)
  name = name or ("Picture " .. pic_id)
  alt = alt or ""

  return xml.elem("w:drawing", nil, {
    xml.elem("wp:inline", {
      ["distT"] = "0",
      ["distB"] = "0",
      ["distL"] = "0",
      ["distR"] = "0",
    }, {
      xml.elem("wp:extent", { ["cx"] = cx, ["cy"] = cy }),
      xml.elem("wp:effectExtent", { ["l"] = "0", ["t"] = "0", ["r"] = "0", ["b"] = "0" }),
      xml.elem("wp:docPr", { ["id"] = pic_id, ["name"] = name, ["descr"] = alt }),
      xml.elem("wp:cNvGraphicFramePr", nil, {
        xml.elem("a:graphicFrameLocks", { ["xmlns:a"] = NS_A, ["noChangeAspect"] = "1" }),
      }),
      xml.elem("a:graphic", { ["xmlns:a"] = NS_A }, {
        xml.elem("a:graphicData", { ["uri"] = NS_PIC }, {
          xml.elem("pic:pic", { ["xmlns:pic"] = NS_PIC }, {
            xml.elem("pic:nvPicPr", nil, {
              xml.elem("pic:cNvPr", { ["id"] = "0", ["name"] = name, ["descr"] = alt }),
              xml.elem("pic:cNvPicPr"),
            }),
            xml.elem("pic:blipFill", nil, {
              xml.elem("a:blip", { ["r:embed"] = r_id }),
              xml.elem("a:stretch", nil, {
                xml.elem("a:fillRect"),
              }),
            }),
            xml.elem("pic:spPr", nil, {
              xml.elem("a:xfrm", nil, {
                xml.elem("a:off", { ["x"] = "0", ["y"] = "0" }),
                xml.elem("a:ext", { ["cx"] = cx, ["cy"] = cy }),
              }),
              xml.elem("a:prstGeom", { ["prst"] = "rect" }, {
                xml.elem("a:avLst"),
              }),
            }),
          }),
        }),
      }),
    }),
  })
end

function M.ast_to_docx(doc, workspace, reference_archive, opts)
  opts = opts or {}
  local preserve = (opts.preserve_styles ~= false) and (reference_archive ~= nil)

  local archive = {
    files = {},
    order = {},
  }

  local preserved_sect_pr = nil
  local preserved_doc_attrs = nil
  local kept_rels = {}
  local max_rid_num = 0
  local header_footer_media = {}

  if preserve and reference_archive and reference_archive.files then
    archive.raw_entries = reference_archive.raw_entries
    for k, v in pairs(reference_archive.files) do
      -- Copy all files except body-specific media that we will re-populate or retain
      if not k:match("^word/media/") then
        archive.files[k] = v
      end
    end
    for _, name in ipairs(reference_archive.order or {}) do
      if archive.files[name] ~= nil then
        archive.order[#archive.order + 1] = name
      end
    end

    -- Inspect original word/document.xml for root attributes and <w:sectPr>
    if reference_archive.files["word/document.xml"] then
      local orig_doc = xml.parse(reference_archive.files["word/document.xml"])
      if orig_doc then
        preserved_doc_attrs = vim.deepcopy(orig_doc.attrs or {})
        local orig_body = xml.find_child(orig_doc, "body")
        if orig_body then
          local sect = xml.find_child(orig_body, "sectPr")
          if sect then
            preserved_sect_pr = sect
          end
        end
      end
    end

    -- Check header/footer rels so their media is always preserved
    for path, content in pairs(reference_archive.files) do
      if path:match("^word/_rels/header.*%.rels$") or path:match("^word/_rels/footer.*%.rels$") then
        local hrels = parse_relationships(reference_archive, path)
        for _, hr in pairs(hrels) do
          if hr.target and hr.target ~= "" then
            local mpath = normalize_zip_path("word", hr.target)
            if reference_archive.files[mpath] then
              header_footer_media[mpath] = reference_archive.files[mpath]
            end
          end
        end
      end
    end
    for mpath, mbytes in pairs(header_footer_media) do
      archive.files[mpath] = mbytes
      archive.order[#archive.order + 1] = mpath
    end

    -- Parse existing word/_rels/document.xml.rels
    local orig_rels = parse_relationships(reference_archive, "word/_rels/document.xml.rels")
    for _, rel in pairs(orig_rels) do
      local rnum = tonumber((rel.id or ""):match("^rId(%d+)$"))
      if rnum and rnum > max_rid_num then
        max_rid_num = rnum
      end
      if rel.type ~= REL_IMAGE and rel.type ~= REL_HYPERLINK and rel.type ~= REL_NUMBERING and rel.type ~= REL_FOOTNOTES then
        kept_rels[#kept_rels + 1] = rel
      end
    end
  end

  local function next_rid()
    max_rid_num = max_rid_num + 1
    return "rId" .. tostring(max_rid_num)
  end

  local function add_rel(rtype, target, target_mode)
    for _, r in ipairs(kept_rels) do
      if r.type == rtype and r.target == target and (r.target_mode or "") == (target_mode or "") then
        return r.id
      end
    end
    local rid = next_rid()
    kept_rels[#kept_rels + 1] = {
      id = rid,
      type = rtype,
      target = target,
      target_mode = target_mode,
    }
    return rid
  end

  -- Ensure styles.xml exists and has required styles
  local styles_xml, heading_style_ids = ensure_required_styles(archive.files["word/styles.xml"])
  archive.files["word/styles.xml"] = styles_xml
  add_rel(REL_STYLES, "styles.xml")

  -- Track media files and drawing IDs
  local pic_counter = 0
  local media_name_used = {}
  local media_exts = {}

  local function resolve_and_embed_image(img_node)
    local src = img_node.src or ""
    if src == "" then
      return nil, nil, nil
    end

    if src:match("^https?://") then
      local rid = add_rel(REL_IMAGE, src, "External")
      local cx, cy = image.resolve_dimensions_emu(nil, img_node.width, img_node.height)
      return rid, cx, cy
    end

    -- Locate image bytes from workspace, filesystem, or reference_archive
    local img_bytes = nil
    local candidates = {}
    if workspace and workspace ~= "" then
      candidates[#candidates + 1] = workspace .. "/" .. src
    end
    candidates[#candidates + 1] = src
    if opts.source_dir and opts.source_dir ~= "" then
      candidates[#candidates + 1] = opts.source_dir .. "/" .. src
    end

    for _, cand in ipairs(candidates) do
      img_bytes = read_binary_file(cand)
      if img_bytes then break end
    end

    if not img_bytes and reference_archive and reference_archive.files then
      local norm = normalize_zip_path("word", src)
      img_bytes = reference_archive.files[norm] or reference_archive.files[src]
    end

    if not img_bytes then
      local rid = add_rel(REL_IMAGE, src, "External")
      local cx, cy = image.resolve_dimensions_emu(nil, img_node.width, img_node.height)
      return rid, cx, cy
    end

    local base_name = vim.fn.fnamemodify(src, ":t")
    if base_name == "" then base_name = "image.png" end
    local ext = vim.fn.fnamemodify(base_name, ":e"):lower()
    if ext == "" then
      ext = "png"
      base_name = base_name .. ".png"
    end
    media_exts[ext] = image.content_type_for_ext(ext)

    local target_rel = "media/" .. base_name
    local zip_media_path = "word/" .. target_rel
    if media_name_used[zip_media_path] and archive.files[zip_media_path] ~= img_bytes then
      local stem = vim.fn.fnamemodify(base_name, ":r")
      local idx = 2
      while true do
        target_rel = string.format("media/%s_%d.%s", stem, idx, ext)
        zip_media_path = "word/" .. target_rel
        if not media_name_used[zip_media_path] or archive.files[zip_media_path] == img_bytes then
          break
        end
        idx = idx + 1
      end
    end

    media_name_used[zip_media_path] = true
    archive.files[zip_media_path] = img_bytes

    local rid = add_rel(REL_IMAGE, target_rel)
    local cx, cy = image.resolve_dimensions_emu(img_bytes, img_node.width, img_node.height)
    return rid, cx, cy
  end

  -- Footnote tracking
  local fn_def_by_id = {}
  for _, fn in ipairs((doc and doc.footnotes) or {}) do
    fn_def_by_id[tostring(fn.id)] = fn.inlines or {}
  end
  local compiled_footnotes = {}
  local fn_id_to_num = {}

  -- Convert AST inlines to WordprocessingML child elements (<w:r>, <w:hyperlink>, etc.)
  local function compile_inlines(inlines, extra_rpr)
    local elements = {}
    for _, item in ipairs(inlines or {}) do
      if item.type == "text" then
        local rpr_children = {}
        if item.code then
          rpr_children[#rpr_children + 1] = xml.elem("w:rStyle", { ["w:val"] = "VerbatimChar" })
          rpr_children[#rpr_children + 1] = xml.elem("w:rFonts", {
            ["w:ascii"] = "Consolas",
            ["w:hAnsi"] = "Consolas",
            ["w:cs"] = "Consolas",
          })
        end
        if extra_rpr and extra_rpr.rstyle and not item.code then
          rpr_children[#rpr_children + 1] = xml.elem("w:rStyle", { ["w:val"] = extra_rpr.rstyle })
        end
        if item.bold or (extra_rpr and extra_rpr.bold) then
          rpr_children[#rpr_children + 1] = xml.elem("w:b")
          rpr_children[#rpr_children + 1] = xml.elem("w:bCs")
        end
        if item.italic or (extra_rpr and extra_rpr.italic) then
          rpr_children[#rpr_children + 1] = xml.elem("w:i")
          rpr_children[#rpr_children + 1] = xml.elem("w:iCs")
        end
        if item.strike then
          rpr_children[#rpr_children + 1] = xml.elem("w:strike")
        end
        if item.underline or (extra_rpr and extra_rpr.underline) then
          rpr_children[#rpr_children + 1] = xml.elem("w:u", { ["w:val"] = "single" })
        end
        if item.sup then
          rpr_children[#rpr_children + 1] = xml.elem("w:vertAlign", { ["w:val"] = "superscript" })
        elseif item.sub then
          rpr_children[#rpr_children + 1] = xml.elem("w:vertAlign", { ["w:val"] = "subscript" })
        end

        local run_children = {}
        if #rpr_children > 0 then
          run_children[#run_children + 1] = xml.elem("w:rPr", nil, rpr_children)
        end
        run_children[#run_children + 1] = xml.elem("w:t", { ["xml:space"] = "preserve" }, {
          xml.text_node(item.text or ""),
        })
        elements[#elements + 1] = xml.elem("w:r", nil, run_children)

      elseif item.type == "math" then
        elements[#elements + 1] = xml.elem("m:oMath", nil, {
          xml.elem("m:r", nil, {
            xml.elem("m:t", nil, { xml.text_node(item.text or "") }),
          }),
        })

      elseif item.type == "linebreak" then
        elements[#elements + 1] = xml.elem("w:r", nil, { xml.elem("w:br") })

      elseif item.type == "tab" then
        elements[#elements + 1] = xml.elem("w:r", nil, { xml.elem("w:tab") })

      elseif item.type == "link" then
        local url = item.url or ""
        local link_runs = compile_inlines(item.inlines or {}, { rstyle = "Hyperlink" })
        if url:sub(1, 1) == "#" then
          elements[#elements + 1] = xml.elem("w:hyperlink", { ["w:anchor"] = url:sub(2) }, link_runs)
        elseif url ~= "" then
          local rid = add_rel(REL_HYPERLINK, url, "External")
          elements[#elements + 1] = xml.elem("w:hyperlink", { ["r:id"] = rid }, link_runs)
        else
          for _, lr in ipairs(link_runs) do
            elements[#elements + 1] = lr
          end
        end

      elseif item.type == "image" then
        local rid, cx, cy = resolve_and_embed_image(item)
        if rid then
          pic_counter = pic_counter + 1
          local pic_name = vim.fn.fnamemodify(item.src or ("Picture " .. pic_counter), ":t")
          local drawing = build_drawing_element(rid, pic_counter, pic_name, item.alt or "", cx, cy)
          elements[#elements + 1] = xml.elem("w:r", nil, {
            xml.elem("w:rPr", nil, { xml.elem("w:noProof") }),
            drawing,
          })
        else
          -- Fallback text if image file does not exist
          elements[#elements + 1] = xml.elem("w:r", nil, {
            xml.elem("w:t", { ["xml:space"] = "preserve" }, {
              xml.text_node(string.format("![%s](%s)", item.alt or "", item.src or "")),
            }),
          })
        end

      elseif item.type == "footnote_ref" then
        local fn_key = tostring(item.id or "1")
        local fn_num = fn_id_to_num[fn_key]
        if not fn_num then
          fn_num = #compiled_footnotes + 1
          fn_id_to_num[fn_key] = fn_num
          compiled_footnotes[#compiled_footnotes + 1] = {
            num = fn_num,
            inlines = fn_def_by_id[fn_key] or { { type = "text", text = fn_key } },
          }
        end
        elements[#elements + 1] = xml.elem("w:r", nil, {
          xml.elem("w:rPr", nil, {
            xml.elem("w:rStyle", { ["w:val"] = "FootnoteReference" }),
          }),
          xml.elem("w:footnoteReference", { ["w:id"] = tostring(fn_num) }),
        })
      end
    end
    return elements
  end

  local existing_numbering_xml = archive.files["word/numbering.xml"]
  local list_id_offset = 0
  if existing_numbering_xml and existing_numbering_xml ~= "" then
    list_id_offset = 1000
    local ex_num_root = xml.parse(existing_numbering_xml)
    if ex_num_root then
      for _, c in ipairs(ex_num_root.children or {}) do
        local aid = tonumber(xml.get_attr(c, "w:abstractNumId"))
        local nid = tonumber(xml.get_attr(c, "w:numId"))
        if aid and aid > list_id_offset then list_id_offset = aid end
        if nid and nid > list_id_offset then list_id_offset = nid end
      end
    end
  end

  local body_children = {}
  local list_defs = {}
  local bullet_num_id = nil
  local active_ordered_num_id = nil
  local bookmark_counter = 0

  for i, block in ipairs((doc and doc.blocks) or {}) do
    local btype = block.type
    local prev_block = doc.blocks[i - 1]

    if btype ~= "list_item" then
      active_ordered_num_id = nil
    end

    if btype == "heading" then
      local lvl = math.min(math.max(tonumber(block.level) or 1, 1), 6)
      local style_id = heading_style_ids[lvl] or ("Heading" .. tostring(lvl))
      bookmark_counter = bookmark_counter + 1
      local bm_id = tostring(bookmark_counter)
      local raw_heading = markdown.render_inlines(block.inlines or {})
      local slug = raw_heading:lower():gsub("[^%w%s%-_]", ""):gsub("%s+", "-"):gsub("^%-+", ""):gsub("%-+$", "")
      if slug == "" then
        slug = "heading-" .. bm_id
      end
      local p_children = {
        xml.elem("w:pPr", nil, {
          xml.elem("w:pStyle", { ["w:val"] = style_id }),
        }),
        xml.elem("w:bookmarkStart", { ["w:id"] = bm_id, ["w:name"] = slug }),
      }
      for _, el in ipairs(compile_inlines(block.inlines)) do
        p_children[#p_children + 1] = el
      end
      p_children[#p_children + 1] = xml.elem("w:bookmarkEnd", { ["w:id"] = bm_id })
      body_children[#body_children + 1] = xml.elem("w:p", nil, p_children)

    elseif btype == "paragraph" then
      local p_children = compile_inlines(block.inlines)
      body_children[#body_children + 1] = xml.elem("w:p", nil, p_children)

    elseif btype == "math_block" then
      body_children[#body_children + 1] = xml.elem("w:p", nil, {
        xml.elem("m:oMathPara", nil, {
          xml.elem("m:oMath", nil, {
            xml.elem("m:r", nil, {
              xml.elem("m:t", nil, { xml.text_node(block.text or "") }),
            }),
          }),
        }),
      })

    elseif btype == "list_item" then
      local lvl = math.min(math.max(tonumber(block.level) or 0, 0), 8)
      local num_id
      if block.ordered then
        if not active_ordered_num_id or (prev_block and prev_block.type == "list_item" and not prev_block.ordered and (prev_block.level or 0) <= lvl) then
          local new_id = list_id_offset + #list_defs + 1
          list_defs[#list_defs + 1] = {
            num_id = new_id,
            abstract_id = new_id,
            ordered = true,
            start = block.number or 1,
          }
          active_ordered_num_id = new_id
        end
        num_id = active_ordered_num_id
      else
        if not bullet_num_id then
          local new_id = list_id_offset + #list_defs + 1
          list_defs[#list_defs + 1] = {
            num_id = new_id,
            abstract_id = new_id,
            ordered = false,
            start = 1,
          }
          bullet_num_id = new_id
        end
        num_id = bullet_num_id
      end

      local item_inlines = vim.deepcopy(block.inlines or {})
      if block.checked == true then
        table.insert(item_inlines, 1, { type = "text", text = "\226\152\145 " })
      elseif block.checked == false then
        table.insert(item_inlines, 1, { type = "text", text = "\226\152\144 " })
      end

      local p_children = {
        xml.elem("w:pPr", nil, {
          xml.elem("w:pStyle", { ["w:val"] = "Compact" }),
          xml.elem("w:numPr", nil, {
            xml.elem("w:ilvl", { ["w:val"] = tostring(lvl) }),
            xml.elem("w:numId", { ["w:val"] = tostring(num_id) }),
          }),
        }),
      }
      for _, el in ipairs(compile_inlines(item_inlines)) do
        p_children[#p_children + 1] = el
      end
      body_children[#body_children + 1] = xml.elem("w:p", nil, p_children)

    elseif btype == "blockquote" then
      local bq_lvl = math.max(tonumber(block.level) or 1, 1)
      local left_ind = tostring(720 * bq_lvl)
      local p_children = {
        xml.elem("w:pPr", nil, {
          xml.elem("w:pStyle", { ["w:val"] = "BlockText" }),
          xml.elem("w:pBdr", nil, {
            xml.elem("w:left", { ["w:val"] = "single", ["w:sz"] = "12", ["w:space"] = "12", ["w:color"] = "D0D7DE" }),
          }),
          xml.elem("w:ind", { ["w:left"] = left_ind, ["w:right"] = "720" }),
        }),
      }
      for _, el in ipairs(compile_inlines(block.inlines)) do
        p_children[#p_children + 1] = el
      end
      body_children[#body_children + 1] = xml.elem("w:p", nil, p_children)

    elseif btype == "code_block" then
      local code_lines = vim.split(block.text or "", "\n", { plain = true })
      for _, cl in ipairs(code_lines) do
        body_children[#body_children + 1] = xml.elem("w:p", nil, {
          xml.elem("w:pPr", nil, {
            xml.elem("w:pStyle", { ["w:val"] = "SourceCode" }),
          }),
          xml.elem("w:r", nil, {
            xml.elem("w:rPr", nil, {
              xml.elem("w:rFonts", { ["w:ascii"] = "Consolas", ["w:hAnsi"] = "Consolas", ["w:cs"] = "Consolas" }),
            }),
            xml.elem("w:t", { ["xml:space"] = "preserve" }, {
              xml.text_node(cl),
            }),
          }),
        })
      end

    elseif btype == "hr" then
      body_children[#body_children + 1] = xml.elem("w:p", nil, {
        xml.elem("w:pPr", nil, {
          xml.elem("w:pBdr", nil, {
            xml.elem("w:bottom", { ["w:val"] = "single", ["w:sz"] = "6", ["w:space"] = "1", ["w:color"] = "auto" }),
          }),
        }),
      })

    elseif btype == "table" then
      local headers = block.headers or {}
      local rows = block.rows or {}
      local alignments = block.alignments or {}
      local num_cols = #headers
      for _, r in ipairs(rows) do
        if #r > num_cols then num_cols = #r end
      end
      if num_cols == 0 then num_cols = 1 end
      local col_w = tostring(math.floor(9350 / num_cols))

      local grid_cols = {}
      for _ = 1, num_cols do
        grid_cols[#grid_cols + 1] = xml.elem("w:gridCol", { ["w:w"] = col_w })
      end

      local tbl_children = {
        xml.elem("w:tblPr", nil, {
          xml.elem("w:tblStyle", { ["w:val"] = "TableGrid" }),
          xml.elem("w:tblW", { ["w:w"] = "0", ["w:type"] = "auto" }),
          xml.elem("w:tblBorders", nil, {
            xml.elem("w:top", { ["w:val"] = "single", ["w:sz"] = "4", ["w:space"] = "0", ["w:color"] = "D0D7DE" }),
            xml.elem("w:left", { ["w:val"] = "single", ["w:sz"] = "4", ["w:space"] = "0", ["w:color"] = "D0D7DE" }),
            xml.elem("w:bottom", { ["w:val"] = "single", ["w:sz"] = "4", ["w:space"] = "0", ["w:color"] = "D0D7DE" }),
            xml.elem("w:right", { ["w:val"] = "single", ["w:sz"] = "4", ["w:space"] = "0", ["w:color"] = "D0D7DE" }),
            xml.elem("w:insideH", { ["w:val"] = "single", ["w:sz"] = "4", ["w:space"] = "0", ["w:color"] = "D0D7DE" }),
            xml.elem("w:insideV", { ["w:val"] = "single", ["w:sz"] = "4", ["w:space"] = "0", ["w:color"] = "D0D7DE" }),
          }),
          xml.elem("w:tblLook", {
            ["w:val"] = "04A0",
            ["w:firstRow"] = "1",
            ["w:lastRow"] = "0",
            ["w:firstColumn"] = "0",
            ["w:lastColumn"] = "0",
            ["w:noHBand"] = "0",
            ["w:noVBand"] = "1",
          }),
        }),
        xml.elem("w:tblGrid", nil, grid_cols),
      }

      local function build_tr(cell_list, is_header)
        local tr_children = {}
        if is_header then
          tr_children[#tr_children + 1] = xml.elem("w:trPr", nil, {
            xml.elem("w:tblHeader", { ["w:val"] = "on" }),
          })
        end
        for c = 1, num_cols do
          local align = alignments[c] or "left"
          local p_children = {
            xml.elem("w:pPr", nil, {
              xml.elem("w:pStyle", { ["w:val"] = "Compact" }),
              xml.elem("w:jc", { ["w:val"] = align }),
            }),
          }
          local cell_inlines = cell_list[c] or {}
          local cell_runs = compile_inlines(cell_inlines)
          for _, cr in ipairs(cell_runs) do
            p_children[#p_children + 1] = cr
          end
          tr_children[#tr_children + 1] = xml.elem("w:tc", nil, {
            xml.elem("w:tcPr", nil, {
              xml.elem("w:tcW", { ["w:w"] = col_w, ["w:type"] = "dxa" }),
            }),
            xml.elem("w:p", nil, p_children),
          })
        end
        return xml.elem("w:tr", nil, tr_children)
      end

      tbl_children[#tbl_children + 1] = build_tr(headers, true)
      for _, r in ipairs(rows) do
        tbl_children[#tbl_children + 1] = build_tr(r, false)
      end

      body_children[#body_children + 1] = xml.elem("w:tbl", nil, tbl_children)
    end
  end

  if #body_children == 0 then
    body_children[1] = xml.elem("w:p")
  end

  -- Append <w:sectPr>
  if preserved_sect_pr then
    body_children[#body_children + 1] = preserved_sect_pr
  else
    body_children[#body_children + 1] = xml.elem("w:sectPr", nil, {
      xml.elem("w:pgSz", { ["w:w"] = "12240", ["w:h"] = "15840" }),
      xml.elem("w:pgMar", {
        ["w:top"] = "1440",
        ["w:right"] = "1440",
        ["w:bottom"] = "1440",
        ["w:left"] = "1440",
        ["w:header"] = "720",
        ["w:footer"] = "720",
        ["w:gutter"] = "0",
      }),
      xml.elem("w:cols", { ["w:space"] = "720" }),
      xml.elem("w:docGrid", { ["w:linePitch"] = "360" }),
    })
  end

  -- Write word/numbering.xml if lists were used
  if #list_defs > 0 then
    archive.files["word/numbering.xml"] = build_numbering_xml(list_defs, existing_numbering_xml)
    add_rel(REL_NUMBERING, "numbering.xml")
  elseif archive.files["word/numbering.xml"] then
    add_rel(REL_NUMBERING, "numbering.xml")
  end

  -- Write word/footnotes.xml if footnotes were used
  if #compiled_footnotes > 0 then
    local fn_children = {
      xml.elem("w:footnote", { ["w:type"] = "separator", ["w:id"] = "-1" }, {
        xml.elem("w:p", nil, { xml.elem("w:r", nil, { xml.elem("w:separator") }) }),
      }),
      xml.elem("w:footnote", { ["w:type"] = "continuationSeparator", ["w:id"] = "0" }, {
        xml.elem("w:p", nil, { xml.elem("w:r", nil, { xml.elem("w:continuationSeparator") }) }),
      }),
    }
    for _, fn in ipairs(compiled_footnotes) do
      local p_children = {
        xml.elem("w:pPr", nil, {
          xml.elem("w:pStyle", { ["w:val"] = "FootnoteText" }),
        }),
        xml.elem("w:r", nil, {
          xml.elem("w:rPr", nil, { xml.elem("w:rStyle", { ["w:val"] = "FootnoteReference" }) }),
          xml.elem("w:footnoteRef"),
        }),
        xml.elem("w:r", nil, {
          xml.elem("w:t", { ["xml:space"] = "preserve" }, { xml.text_node(" ") }),
        }),
      }
      for _, r in ipairs(compile_inlines(fn.inlines)) do
        p_children[#p_children + 1] = r
      end
      fn_children[#fn_children + 1] = xml.elem("w:footnote", { ["w:id"] = tostring(fn.num) }, {
        xml.elem("w:p", nil, p_children),
      })
    end
    local fn_root = xml.elem("w:footnotes", { ["xmlns:w"] = NS_W, ["xmlns:r"] = NS_R }, fn_children)
    archive.files["word/footnotes.xml"] = xml.serialize(fn_root)
    add_rel(REL_FOOTNOTES, "footnotes.xml")
  elseif archive.files["word/footnotes.xml"] then
    add_rel(REL_FOOTNOTES, "footnotes.xml")
  end

  -- Ensure settings.xml, webSettings.xml, fontTable.xml exist
  if not archive.files["word/settings.xml"] then
    archive.files["word/settings.xml"] = xml.serialize(xml.elem("w:settings", { ["xmlns:w"] = NS_W }, {
      xml.elem("w:zoom", { ["w:percent"] = "100" }),
      xml.elem("w:defaultTabStop", { ["w:val"] = "720" }),
    }))
  end
  add_rel(REL_SETTINGS, "settings.xml")

  if not archive.files["word/webSettings.xml"] then
    archive.files["word/webSettings.xml"] = xml.serialize(xml.elem("w:webSettings", { ["xmlns:w"] = NS_W }, {
      xml.elem("w:optimizeForBrowser"),
    }))
  end
  add_rel(REL_WEB_SETTINGS, "webSettings.xml")

  if not archive.files["word/fontTable.xml"] then
    archive.files["word/fontTable.xml"] = xml.serialize(xml.elem("w:fonts", { ["xmlns:w"] = NS_W }, {
      xml.elem("w:font", { ["w:name"] = "Calibri" }),
      xml.elem("w:font", { ["w:name"] = "Consolas" }),
      xml.elem("w:font", { ["w:name"] = "Symbol" }),
    }))
  end
  add_rel(REL_FONT_TABLE, "fontTable.xml")

  -- Build word/document.xml
  local doc_attrs = preserved_doc_attrs or {}
  doc_attrs["xmlns:w"] = NS_W
  doc_attrs["xmlns:r"] = NS_R
  doc_attrs["xmlns:wp"] = NS_WP
  doc_attrs["xmlns:a"] = NS_A
  doc_attrs["xmlns:pic"] = NS_PIC
  doc_attrs["xmlns:v"] = NS_V
  doc_attrs["xmlns:o"] = NS_O
  doc_attrs["xmlns:m"] = NS_M

  local doc_elem = xml.elem("w:document", doc_attrs, {
    xml.elem("w:body", nil, body_children),
  })
  archive.files["word/document.xml"] = xml.serialize(doc_elem)

  -- Build word/_rels/document.xml.rels
  local rel_elems = {}
  for _, r in ipairs(kept_rels) do
    local rattrs = {
      ["Id"] = r.id,
      ["Type"] = r.type,
      ["Target"] = r.target,
    }
    if r.target_mode and r.target_mode ~= "" then
      rattrs["TargetMode"] = r.target_mode
    end
    rel_elems[#rel_elems + 1] = xml.elem("Relationship", rattrs)
  end
  archive.files["word/_rels/document.xml.rels"] = xml.serialize(
    xml.elem("Relationships", { ["xmlns"] = NS_PKG_REL }, rel_elems)
  )

  -- Ensure _rels/.rels exists
  if not archive.files["_rels/.rels"] then
    archive.files["_rels/.rels"] = xml.serialize(
      xml.elem("Relationships", { ["xmlns"] = NS_PKG_REL }, {
        xml.elem("Relationship", { ["Id"] = "rId1", ["Type"] = REL_OFFICE_DOC, ["Target"] = "word/document.xml" }),
        xml.elem("Relationship", { ["Id"] = "rId2", ["Type"] = REL_CORE_PROPS, ["Target"] = "docProps/core.xml" }),
        xml.elem("Relationship", { ["Id"] = "rId3", ["Type"] = REL_EXT_PROPS, ["Target"] = "docProps/app.xml" }),
      })
    )
  end

  -- Ensure docProps/core.xml and docProps/app.xml exist (and apply metadata if provided)
  local meta = (doc and doc.metadata) or {}
  if not archive.files["docProps/core.xml"] or next(meta) ~= nil then
    local core_children = {
      xml.elem("dc:title", nil, meta.title and { xml.text_node(meta.title) } or nil),
      xml.elem("dc:subject", nil, meta.subject and { xml.text_node(meta.subject) } or nil),
      xml.elem("dc:creator", nil, { xml.text_node(meta.author or meta.creator or "nvim-docx") }),
      xml.elem("cp:keywords", nil, meta.keywords and { xml.text_node(meta.keywords) } or nil),
      xml.elem("dc:description", nil, meta.description and { xml.text_node(meta.description) } or nil),
      xml.elem("cp:revision", nil, { xml.text_node("1") }),
    }
    archive.files["docProps/core.xml"] = xml.serialize(xml.elem("cp:coreProperties", {
      ["xmlns:cp"] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
      ["xmlns:dc"] = "http://purl.org/dc/elements/1.1/",
      ["xmlns:dcterms"] = "http://purl.org/dc/terms/",
      ["xmlns:dcmitype"] = "http://purl.org/dc/dcmitype/",
      ["xmlns:xsi"] = "http://www.w3.org/2001/XMLSchema-instance",
    }, core_children))
  end

  if not archive.files["docProps/app.xml"] then
    archive.files["docProps/app.xml"] = xml.serialize(xml.elem("Properties", {
      ["xmlns"] = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
      ["xmlns:vt"] = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    }, {
      xml.elem("Application", nil, { xml.text_node("nvim-docx") }),
    }))
  end

  -- Build / update [Content_Types].xml
  local ct_root = archive.files["[Content_Types].xml"] and xml.parse(archive.files["[Content_Types].xml"]) or nil
  if not ct_root then
    ct_root = xml.elem("Types", { ["xmlns"] = NS_CT }, {
      xml.elem("Default", { ["Extension"] = "rels", ["ContentType"] = "application/vnd.openxmlformats-package.relationships+xml" }),
      xml.elem("Default", { ["Extension"] = "xml", ["ContentType"] = "application/xml" }),
    })
  end

  local existing_defaults = {}
  local existing_overrides = {}
  for _, child in ipairs(ct_root.children or {}) do
    if child.local_name == "Default" then
      local ext = (xml.get_attr(child, "Extension") or ""):lower()
      existing_defaults[ext] = true
    elseif child.local_name == "Override" then
      local pn = xml.get_attr(child, "PartName") or ""
      existing_overrides[pn] = true
    end
  end

  for ext, ctype in pairs(media_exts) do
    if not existing_defaults[ext] then
      ct_root.children[#ct_root.children + 1] = xml.elem("Default", {
        ["Extension"] = ext,
        ["ContentType"] = ctype,
      })
      existing_defaults[ext] = true
    end
  end

  local required_overrides = {
    { "/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" },
    { "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" },
    { "/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" },
    { "/word/webSettings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml" },
    { "/word/fontTable.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" },
    { "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml" },
    { "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml" },
  }
  if archive.files["word/numbering.xml"] then
    required_overrides[#required_overrides + 1] = {
      "/word/numbering.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    }
  end
  if archive.files["word/footnotes.xml"] then
    required_overrides[#required_overrides + 1] = {
      "/word/footnotes.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    }
  end
  for _, ov in ipairs(required_overrides) do
    if not existing_overrides[ov[1]] then
      ct_root.children[#ct_root.children + 1] = xml.elem("Override", {
        ["PartName"] = ov[1],
        ["ContentType"] = ov[2],
      })
      existing_overrides[ov[1]] = true
    end
  end

  archive.files["[Content_Types].xml"] = xml.serialize(ct_root)

  -- Ensure [Content_Types].xml and _rels/.rels are first in ZIP order
  local base_order = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/_rels/document.xml.rels",
    "word/document.xml",
  }
  for _, b in ipairs(archive.order) do
    base_order[#base_order + 1] = b
  end
  archive.order = base_order

  return archive
end

function M.get_docx_stats(archive, ast)
  ast = ast or M.docx_to_ast(archive)
  local stats = {
    metadata = (ast and ast.metadata) or {},
    headings = 0,
    paragraphs = 0,
    list_items = 0,
    tables = 0,
    code_blocks = 0,
    blockquotes = 0,
    images = 0,
    footnotes = (ast and ast.footnotes and #ast.footnotes) or 0,
    words = 0,
    chars = 0,
    archive_files = 0,
  }

  if archive and archive.files then
    for name, _ in pairs(archive.files) do
      stats.archive_files = stats.archive_files + 1
      if name:match("^word/media/") then
        stats.images = stats.images + 1
      end
    end
  end

  local function count_inlines(inlines)
    for _, item in ipairs(inlines or {}) do
      if (item.type == "text" or item.type == "math") and item.text then
        stats.chars = stats.chars + #item.text
        for _ in item.text:gmatch("%S+") do
          stats.words = stats.words + 1
        end
      elseif item.type == "link" and item.inlines then
        count_inlines(item.inlines)
      end
    end
  end

  for _, block in ipairs((ast and ast.blocks) or {}) do
    local bt = block.type
    if bt == "heading" then
      stats.headings = stats.headings + 1
      count_inlines(block.inlines)
    elseif bt == "paragraph" then
      stats.paragraphs = stats.paragraphs + 1
      count_inlines(block.inlines)
    elseif bt == "list_item" then
      stats.list_items = stats.list_items + 1
      count_inlines(block.inlines)
    elseif bt == "blockquote" then
      stats.blockquotes = stats.blockquotes + 1
      count_inlines(block.inlines)
    elseif bt == "code_block" or bt == "math_block" then
      if bt == "code_block" then
        stats.code_blocks = stats.code_blocks + 1
      else
        stats.paragraphs = stats.paragraphs + 1
      end
      local txt = block.text or ""
      stats.chars = stats.chars + #txt
      for _ in txt:gmatch("%S+") do
        stats.words = stats.words + 1
      end
    elseif bt == "table" then
      stats.tables = stats.tables + 1
      for _, h in ipairs(block.headers or {}) do
        count_inlines(h)
      end
      for _, r in ipairs(block.rows or {}) do
        for _, c in ipairs(r) do
          count_inlines(c)
        end
      end
    end
  end

  return stats
end

return M

