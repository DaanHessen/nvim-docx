-- Pure-Lua reader for legacy binary Word documents: [MS-CFB] containers holding [MS-DOC] Word 97-2003
-- (and Word 6/95) files, plus the pre-OLE Word for Windows 1/2 and Write formats. Produces the nvim-docx AST
-- with headings, lists, tables, character formatting, hyperlinks, footnotes/endnotes and pictures.

local image = require("nvim-docx.image")
local markdown = require("nvim-docx.markdown")

local M = {}

local OLE2_MAGIC = "\208\207\17\224\161\177\26\225"

local function u16(s, pos)
  local b1, b2 = string.byte(s, pos, pos + 1)
  if not b2 then return 0 end
  return b1 + b2 * 256
end

local function u32(s, pos)
  local b1, b2, b3, b4 = string.byte(s, pos, pos + 3)
  if not b4 then return 0 end
  return b1 + b2 * 256 + b3 * 65536 + b4 * 16777216
end

local function codepoint_to_utf8(cp)
  if cp < 0 or cp > 0x10FFFF then
    return ""
  elseif cp < 0x80 then
    return string.char(cp)
  elseif cp < 0x800 then
    return string.char(
      0xC0 + math.floor(cp / 64),
      0x80 + (cp % 64)
    )
  elseif cp < 0x10000 then
    return string.char(
      0xE0 + math.floor(cp / 4096),
      0x80 + (math.floor(cp / 64) % 64),
      0x80 + (cp % 64)
    )
  else
    return string.char(
      0xF0 + math.floor(cp / 262144),
      0x80 + (math.floor(cp / 4096) % 64),
      0x80 + (math.floor(cp / 64) % 64),
      0x80 + (cp % 64)
    )
  end
end

-- Windows-1252 (0x80..0x9F) mapping to Unicode codepoints
local CP1252_MAP = {
  [0x80] = 0x20AC, [0x82] = 0x201A, [0x83] = 0x0192, [0x84] = 0x201E,
  [0x85] = 0x2026, [0x86] = 0x2020, [0x87] = 0x2021, [0x88] = 0x02C6,
  [0x89] = 0x2030, [0x8A] = 0x0160, [0x8B] = 0x2039, [0x8C] = 0x0152,
  [0x8E] = 0x017D, [0x91] = 0x2018, [0x92] = 0x2019, [0x93] = 0x201C,
  [0x94] = 0x201D, [0x95] = 0x2022, [0x96] = 0x2013, [0x97] = 0x2014,
  [0x98] = 0x02DC, [0x99] = 0x2122, [0x9A] = 0x0161, [0x9B] = 0x203A,
  [0x9C] = 0x0153, [0x9E] = 0x017E, [0x9F] = 0x0178,
}

local function cp1252_to_utf8(bytes)
  local out = {}
  for i = 1, #bytes do
    local b = string.byte(bytes, i)
    if b < 0x80 then
      out[#out + 1] = string.char(b)
    elseif b <= 0x9F then
      local cp = CP1252_MAP[b]
      out[#out + 1] = cp and codepoint_to_utf8(cp) or " "
    else
      out[#out + 1] = codepoint_to_utf8(b)
    end
  end
  return table.concat(out)
end

local function utf16le_to_utf8(bytes)
  local out = {}
  local i = 1
  local len = #bytes
  while i + 1 <= len do
    local w1 = u16(bytes, i)
    i = i + 2
    if w1 >= 0xD800 and w1 <= 0xDBFF and i + 1 <= len then
      local w2 = u16(bytes, i)
      if w2 >= 0xDC00 and w2 <= 0xDFFF then
        i = i + 2
        local cp = 0x10000 + (w1 - 0xD800) * 1024 + (w2 - 0xDC00)
        out[#out + 1] = codepoint_to_utf8(cp)
      else
        out[#out + 1] = codepoint_to_utf8(w1)
      end
    else
      out[#out + 1] = codepoint_to_utf8(w1)
    end
  end
  return table.concat(out)
end

function M.is_ole2_file(filepath)
  local f = io.open(filepath, "rb")
  if not f then return false end
  local hdr = f:read(8)
  f:close()
  return hdr == OLE2_MAGIC
end

-- Parse OLE2 Compound File Binary streams
local function read_ole2_streams(data)
  if not data or #data < 512 or data:sub(1, 8) ~= OLE2_MAGIC then
    return nil, "Not a valid OLE2 Compound File Binary"
  end

  local sector_shift = u16(data, 31)
  local mini_shift = u16(data, 33)
  local sector_size = 2 ^ sector_shift
  local mini_sector_size = 2 ^ mini_shift
  local num_fat_sectors = u32(data, 45)
  local first_dir_sector = u32(data, 49)
  local mini_cutoff = u32(data, 57)
  local first_minifat_sector = u32(data, 61)
  local num_minifat_sectors = u32(data, 65)
  local first_difat_sector = u32(data, 69)
  local num_difat_sectors = u32(data, 73)

  local fat_sector_ids = {}
  for i = 0, math.min(num_fat_sectors, 109) - 1 do
    local sec = u32(data, 77 + i * 4)
    if sec < 0xFFFFFFFC then
      fat_sector_ids[#fat_sector_ids + 1] = sec
    end
  end

  -- Follow DIFAT chain if > 109 FAT sectors
  local difat_sec = first_difat_sector
  local difat_count = 0
  local entries_per_difat = (sector_size / 4) - 1
  while difat_sec < 0xFFFFFFFC and difat_count < num_difat_sectors and #fat_sector_ids < num_fat_sectors do
    difat_count = difat_count + 1
    local offset = (difat_sec + 1) * sector_size + 1
    if offset + sector_size - 1 > #data then break end
    for j = 0, entries_per_difat - 1 do
      local sec = u32(data, offset + j * 4)
      if sec < 0xFFFFFFFC then
        fat_sector_ids[#fat_sector_ids + 1] = sec
      end
    end
    difat_sec = u32(data, offset + entries_per_difat * 4)
  end

  local fat = {}
  local entries_per_sector = sector_size / 4
  for _, sec in ipairs(fat_sector_ids) do
    local offset = (sec + 1) * sector_size + 1
    if offset + sector_size - 1 <= #data then
      for j = 0, entries_per_sector - 1 do
        fat[#fat + 1] = u32(data, offset + j * 4)
      end
    end
  end

  local function read_sector_chain(start_sec, max_bytes)
    local chunks = {}
    local sec = start_sec
    local total = 0
    local visited = {}
    while sec and sec < 0xFFFFFFFC and not visited[sec] do
      visited[sec] = true
      local offset = (sec + 1) * sector_size + 1
      if offset > #data then break end
      local chunk_end = math.min(#data, offset + sector_size - 1)
      chunks[#chunks + 1] = data:sub(offset, chunk_end)
      total = total + (chunk_end - offset + 1)
      if max_bytes and total >= max_bytes then break end
      sec = fat[sec + 1]
    end
    local combined = table.concat(chunks)
    if max_bytes and #combined > max_bytes then
      return combined:sub(1, max_bytes)
    end
    return combined
  end

  local dir_bytes = read_sector_chain(first_dir_sector)
  local all_entries = {}
  for pos = 1, #dir_bytes - 127, 128 do
    local name_len = u16(dir_bytes, pos + 64)
    local chars = {}
    for k = 0, math.floor(name_len / 2) - 2 do
      local ch = u16(dir_bytes, pos + k * 2)
      if ch > 0 and ch < 128 then
        chars[#chars + 1] = string.char(ch)
      end
    end
    all_entries[#all_entries + 1] = {
      name = table.concat(chars),
      type = string.byte(dir_bytes, pos + 66) or 0,
      left = u32(dir_bytes, pos + 68),
      right = u32(dir_bytes, pos + 72),
      child = u32(dir_bytes, pos + 76),
      start_sec = u32(dir_bytes, pos + 116),
      size = u32(dir_bytes, pos + 120),
    }
  end
  local root_entry = all_entries[1]
  if not root_entry or root_entry.type ~= 5 then
    return nil, "Invalid OLE2 directory (no root entry)"
  end

  -- Only streams directly under the root storage belong to the document; embedded objects
  -- (ObjectPool, OLE packages) carry their own WordDocument/Table streams in sub-storages
  local entries = {}
  local stack, seen = { root_entry.child }, {}
  while #stack > 0 do
    local id = table.remove(stack)
    local e = id < 0xFFFFFFFA and all_entries[id + 1] or nil
    if e and not seen[id] then
      seen[id] = true
      if e.type == 2 then
        entries[e.name] = e
      end
      stack[#stack + 1] = e.left
      stack[#stack + 1] = e.right
    end
  end

  -- Build MiniFAT & MiniStream if needed
  local minifat = {}
  if first_minifat_sector < 0xFFFFFFFC and num_minifat_sectors > 0 then
    local mf_bytes = read_sector_chain(first_minifat_sector, num_minifat_sectors * sector_size)
    for pos = 1, #mf_bytes - 3, 4 do
      minifat[#minifat + 1] = u32(mf_bytes, pos)
    end
  end
  local ministream = ""
  if root_entry and root_entry.start_sec < 0xFFFFFFFC and root_entry.size > 0 then
    ministream = read_sector_chain(root_entry.start_sec, root_entry.size)
  end

  local function read_mini_chain(start_sec, max_bytes)
    local chunks = {}
    local sec = start_sec
    local total = 0
    local visited = {}
    while sec and sec < 0xFFFFFFFC and not visited[sec] do
      visited[sec] = true
      local offset = sec * mini_sector_size + 1
      if offset > #ministream then break end
      local chunk_end = math.min(#ministream, offset + mini_sector_size - 1)
      chunks[#chunks + 1] = ministream:sub(offset, chunk_end)
      total = total + (chunk_end - offset + 1)
      if max_bytes and total >= max_bytes then break end
      sec = minifat[sec + 1]
    end
    local combined = table.concat(chunks)
    if max_bytes and #combined > max_bytes then
      return combined:sub(1, max_bytes)
    end
    return combined
  end

  local streams = {}
  for name, entry in pairs(entries) do
    if entry.size == 0 then
      streams[name] = ""
    elseif entry.size < mini_cutoff and #ministream > 0 then
      streams[name] = read_mini_chain(entry.start_sec, entry.size)
    else
      streams[name] = read_sector_chain(entry.start_sec, entry.size)
    end
  end

  -- Stream names are case-insensitive in compound files
  return setmetatable(streams, {
    __index = function(t, name)
      local lower = type(name) == "string" and name:lower() or nil
      for k, v in pairs(t) do
        if k:lower() == lower then
          return v
        end
      end
    end,
  })
end

--------------------------------------------------------------------------------
-- Format detection
--------------------------------------------------------------------------------

local WORD2_IDENTS = { [0xA59B] = true, [0xA5DB] = true, [0xA59C] = true }
local WRITE_IDENTS = { [0xBE31] = true, [0xBE32] = true }

-- "ole" (Word 6/95/97-2003, possibly encrypted .docx), "word2" (Word for Windows 1/2), "write", or nil
function M.detect_legacy(filepath)
  local f = io.open(filepath, "rb")
  if not f then
    return nil
  end
  local hdr = f:read(8) or ""
  f:close()
  if hdr == OLE2_MAGIC then
    return "ole"
  end
  local ident = u16(hdr, 1)
  if WORD2_IDENTS[ident] then
    return "word2"
  end
  if WRITE_IDENTS[ident] then
    return "write"
  end
  return nil
end

function M.is_legacy_word_file(filepath)
  return M.detect_legacy(filepath) ~= nil
end

-- Windows code page for a Word 6/95 language id
local LID_CODEPAGE = {
  [0x19] = 1251, [0x22] = 1251, [0x23] = 1251, [0x02] = 1251, [0x2F] = 1251,
  [0x05] = 1250, [0x0E] = 1250, [0x15] = 1250, [0x18] = 1250, [0x1B] = 1250, [0x24] = 1250, [0x1C] = 1250,
  [0x08] = 1253, [0x1F] = 1254, [0x0D] = 1255, [0x01] = 1256,
  [0x25] = 1257, [0x26] = 1257, [0x27] = 1257, [0x2A] = 1258, [0x1E] = 874,
}

local function codepage_for_lid(lid)
  local primary = lid % 1024
  if primary == 0x1A then
    -- Serbian/Croatian/Bosnian: Cyrillic sublanguages use 1251
    local sub = math.floor(lid / 1024)
    return (sub == 3 or sub == 7) and 1251 or 1250
  end
  return LID_CODEPAGE[primary] or 1252
end

local function bytes_to_utf8(bytes, codepage)
  if not codepage or codepage == 1252 then
    return cp1252_to_utf8(bytes)
  end
  local upper = require("nvim-docx.codepages")[codepage]
  if not upper then
    return cp1252_to_utf8(bytes)
  end
  local out = {}
  for i = 1, #bytes do
    local b = string.byte(bytes, i)
    if b < 0x80 then
      out[i] = string.char(b)
    else
      local cp = upper[b - 0x7F]
      out[i] = (cp and cp > 0) and codepoint_to_utf8(cp) or "?"
    end
  end
  return table.concat(out)
end

--------------------------------------------------------------------------------
-- Pictures (OfficeArt BLIPs)
--------------------------------------------------------------------------------

local zip = nil
local function inflate_zlib(data)
  zip = zip or require("nvim-docx.zip")
  -- zlib stream: 2-byte header, raw DEFLATE, 4-byte Adler-32
  local ok, out = pcall(zip.inflate, data:sub(3))
  return ok and out or nil
end

-- BLIP record types -> extension, header bytes after the record header (UIDs + tag/metafile header)
local BLIP_TYPES = {
  [0xF01A] = { ext = "emf", meta = true, uid1 = 0x3D4 },
  [0xF01B] = { ext = "wmf", meta = true, uid1 = 0x216 },
  [0xF01C] = { ext = "pict", meta = true, uid1 = 0x542 },
  [0xF01D] = { ext = "jpeg", uid1 = 0x46A, uid1b = 0x6E2 },
  [0xF02A] = { ext = "jpeg", uid1 = 0x46A, uid1b = 0x6E2 },
  [0xF01E] = { ext = "png", uid1 = 0x6E0 },
  [0xF01F] = { ext = "bmp", dib = true, uid1 = 0x7A8 },
  [0xF029] = { ext = "tiff", uid1 = 0x6E4 },
}

-- Decode one BLIP record starting at pos (1-based) in buf; returns bytes, ext
local function decode_blip(buf, pos)
  local ver_inst = u16(buf, pos)
  local rec_type = u16(buf, pos + 2)
  local rec_len = u32(buf, pos + 4)
  local info = BLIP_TYPES[rec_type]
  if not info or pos + 7 + rec_len > #buf then
    return nil
  end
  local inst = math.floor(ver_inst / 16)
  local body = pos + 8
  local uids = (inst == info.uid1 or inst == info.uid1b) and 16 or 32
  if info.meta then
    -- 34-byte metafile header: cbSize at 0, ..., cbSave at 28, compression byte at 32 (0 = DEFLATE)
    local mh = body + uids
    local cb_save = u32(buf, mh + 28)
    local compression = string.byte(buf, mh + 32)
    local data = buf:sub(mh + 34, mh + 33 + cb_save)
    if compression == 0 then
      data = inflate_zlib(data)
    end
    if not data then
      return nil
    end
    if info.ext == "pict" then
      return nil -- Macintosh PICT has no place in .docx
    end
    return data, info.ext
  end
  local data = buf:sub(body + uids + 1, pos + 7 + rec_len)
  if info.dib then
    -- DIB without the 14-byte BITMAPFILEHEADER; add one to make a .bmp file
    local header_size = u32(data, 1)
    local bit_count = u16(data, 15)
    local clr_used = u32(data, 33)
    local palette = 0
    if bit_count <= 8 then
      palette = (clr_used > 0 and clr_used or 2 ^ bit_count) * 4
    end
    local off_bits = 14 + header_size + palette
    local size = 14 + #data
    local function le32(n)
      return string.char(n % 256, math.floor(n / 256) % 256, math.floor(n / 65536) % 256, math.floor(n / 16777216) % 256)
    end
    data = "BM" .. le32(size) .. "\0\0\0\0" .. le32(off_bits) .. data
  end
  return data, info.ext
end

-- Find the first BLIP inside [first, last] of buf (records nest; scan record headers)
local function find_blip(buf, first, last)
  local pos = first
  while pos + 7 <= last do
    local rec_type = u16(buf, pos + 2)
    local rec_len = u32(buf, pos + 4)
    local ver = u16(buf, pos) % 16
    if BLIP_TYPES[rec_type] then
      local data, ext = decode_blip(buf, pos)
      if data then
        return data, ext
      end
    end
    if rec_type == 0xF007 then
      -- BSE: 36-byte header, embedded BLIP follows when present
      local data, ext = find_blip(buf, pos + 8 + 36, math.min(last, pos + 7 + rec_len))
      if data then
        return data, ext
      end
      pos = pos + 8 + rec_len
    elseif ver == 0xF and rec_type >= 0xF000 then
      pos = pos + 8 -- container: descend
    elseif rec_type >= 0xF000 and rec_type <= 0xF2FF and rec_len < last then
      pos = pos + 8 + rec_len
    else
      pos = pos + 1
    end
  end
  return nil
end

--------------------------------------------------------------------------------
-- Property modifiers (SPRMs)
--------------------------------------------------------------------------------

local SPRM = {
  -- character
  CFRMarkDel = 0x0800, CFRMarkIns = 0x0801, CFData = 0x0806, CFOle2 = 0x080A,
  CFBold = 0x0835, CFItalic = 0x0836, CFStrike = 0x0837, CFVanish = 0x083C, CFSpec = 0x0855,
  CFDStrike = 0x2A53, CKul = 0x2A3E, CIss = 0x2A48, CIstd = 0x4A30, CPicLocation = 0x6A03,
  CHps = 0x4A43, CRgFtc0 = 0x4A4F, CSymbol = 0x6A09, CFtcDefault = 0x4A3D, CIco = 0x2A42, CCv = 0x6870,
  -- paragraph
  PIstd = 0x4600, PIlvl = 0x260A, PIlfo = 0x460B, PFInTable = 0x2416, PFTtp = 0x2417,
  PItap = 0x6649, PFInnerTtp = 0x244C, PFInnerTableCell = 0x244B, POutLvl = 0x2640,
  PJc = 0x2461, PJc80 = 0x2403,
}

-- Iterate the SPRMs of a grpprl: fn(sprm, operand_pos, operand_len)
local function each_sprm(buf, pos, last, fn)
  while pos + 1 <= last do
    local sprm = u16(buf, pos)
    local spra = math.floor(sprm / 8192)
    local op = pos + 2
    local len
    if spra == 0 or spra == 1 then
      len = 1
    elseif spra == 2 or spra == 4 or spra == 5 then
      len = 2
    elseif spra == 3 then
      len = 4
    elseif spra == 7 then
      len = 3
    else
      if sprm == 0xD608 or sprm == 0xC615 then
        local cb = u16(buf, op)
        if sprm == 0xC615 and cb == 255 then
          cb = string.byte(buf, op) or 0
          op = op + 1
        else
          op = op + 2
          cb = cb - 1
        end
        len = math.max(cb, 0)
      else
        len = string.byte(buf, op) or 0
        op = op + 1
      end
    end
    if op + len - 1 > last + 1 then
      return
    end
    fn(sprm, op, len)
    pos = op + len
  end
end

-- Apply a character grpprl onto props (toggles: 0 off, 1 on, 0x80 = style value, 0x81 = inverse of style)
-- Boolean character properties may be stored as 1 or as 0x81 ("opposite of the style", which for
-- these is off)
local function flag_on(buf, op)
  local v = string.byte(buf, op)
  return v == 1 or v == 0x81
end

local TOGGLES = {
  [SPRM.CFBold] = "bold", [SPRM.CFItalic] = "italic", [SPRM.CFStrike] = "strike",
  [SPRM.CFDStrike] = "strike", [SPRM.CFVanish] = "vanish",
}
local function apply_chp(props, style, buf, pos, last)
  each_sprm(buf, pos, last, function(sprm, op, len)
    local key = TOGGLES[sprm]
    if key then
      local v = string.byte(buf, op)
      if v == 0x80 then
        props[key] = style and style[key] or false
      elseif v == 0x81 then
        props[key] = not (style and style[key])
      else
        props[key] = v ~= 0
      end
    elseif sprm == SPRM.CKul then
      props.underline = (string.byte(buf, op) or 0) ~= 0
    elseif sprm == SPRM.CIss then
      local v = string.byte(buf, op) or 0
      props.sup = v == 1
      props.sub = v == 2
    elseif sprm == SPRM.CFSpec then
      props.special = flag_on(buf, op)
    elseif sprm == SPRM.CPicLocation then
      props.pic = u32(buf, op)
    elseif sprm == SPRM.CFData then
      props.data = flag_on(buf, op)
    elseif sprm == SPRM.CFOle2 then
      props.ole2 = flag_on(buf, op)
    elseif sprm == SPRM.CFRMarkDel then
      props.deleted = flag_on(buf, op)
    elseif sprm == SPRM.CIstd then
      props.istd = u16(buf, op)
    elseif sprm == SPRM.CRgFtc0 then
      props.ftc = u16(buf, op)
    elseif sprm == SPRM.CSymbol then
      props.sym = { ftc = u16(buf, op), char = u16(buf, op + 2) }
    end
  end)
end

local function apply_pap(props, buf, pos, last)
  each_sprm(buf, pos, last, function(sprm, op, len)
    if sprm == SPRM.PIstd then
      props.istd = u16(buf, op)
    elseif sprm == SPRM.PIlvl then
      props.ilvl = string.byte(buf, op)
    elseif sprm == SPRM.PIlfo then
      props.ilfo = u16(buf, op)
    elseif sprm == SPRM.PFInTable then
      props.in_table = flag_on(buf, op)
    elseif sprm == SPRM.PFTtp then
      props.ttp = flag_on(buf, op)
    elseif sprm == SPRM.PItap then
      props.itap = u32(buf, op)
    elseif sprm == SPRM.PFInnerTtp then
      props.inner_ttp = flag_on(buf, op)
    elseif sprm == SPRM.POutLvl then
      props.outline = string.byte(buf, op)
    end
  end)
end

-- Word 6/95 SPRMs have one-byte opcodes. Operand sizes; false = length-prefixed, missing = unknown
local SPRM6_SIZE = {
  [2] = 2, [3] = false, [4] = 1, [5] = 1, [6] = 1, [7] = 1, [8] = 1, [9] = 1, [10] = 1, [11] = 1,
  [12] = false, [13] = 1, [14] = 1, [15] = false, [16] = 2, [17] = 2, [18] = 2, [19] = 2, [20] = 4,
  [21] = 2, [22] = 2, [23] = false, [24] = 1, [25] = 1, [26] = 2, [27] = 2, [28] = 2, [29] = 1,
  [30] = 2, [31] = 2, [32] = 2, [33] = 2, [34] = 2, [35] = 2, [36] = 2, [37] = 1, [38] = 2, [39] = 2,
  [40] = 2, [41] = 2, [42] = 2, [43] = 2, [44] = 1, [45] = 2, [46] = 2, [47] = 2, [48] = 2, [49] = 2,
  [50] = 1, [51] = 1, [52] = false,
  [65] = 1, [66] = 1, [67] = 1, [68] = false, [69] = 2, [70] = 4, [71] = 1, [72] = 2, [73] = 3,
  [74] = false, [75] = 1, [77] = 1, [80] = 2, [81] = false, [82] = 0, [83] = 0, [85] = 1, [86] = 1,
  [87] = 1, [88] = 1, [89] = 1, [90] = 1, [91] = 1, [92] = 1, [93] = 2, [94] = 1, [95] = 3, [96] = 2,
  [97] = 2, [98] = 1, [99] = 2, [100] = 1, [101] = 2, [102] = 1, [103] = false, [104] = 1,
  [105] = false, [106] = false, [107] = 2, [108] = false, [109] = 2, [110] = 2, [117] = 1, [118] = 1,
}

local function each_sprm6(buf, pos, last, fn)
  while pos <= last do
    local sprm = string.byte(buf, pos)
    local size = SPRM6_SIZE[sprm]
    local op = pos + 1
    local len
    if size == nil then
      return -- unknown opcode: its length is unknown, stop rather than misread
    elseif size == false then
      len = string.byte(buf, op) or 0
      op = op + 1
    else
      len = size
    end
    if op + len - 1 > last then
      return
    end
    fn(sprm, op, len)
    pos = op + len
  end
end

local TOGGLES6 = { [85] = "bold", [86] = "italic", [87] = "strike", [92] = "vanish" }

local function apply_chp6(props, style, buf, pos, last)
  each_sprm6(buf, pos, last, function(sprm, op, len)
    local key = TOGGLES6[sprm]
    if key then
      local v = string.byte(buf, op)
      if v == 0x80 then
        props[key] = style and style[key] or false
      elseif v == 0x81 then
        props[key] = not (style and style[key])
      else
        props[key] = v ~= 0
      end
    elseif sprm == 94 then
      props.underline = (string.byte(buf, op) or 0) ~= 0
    elseif sprm == 104 then
      local v = string.byte(buf, op) or 0
      props.sup = v == 1
      props.sub = v == 2
    elseif sprm == 117 then
      props.special = flag_on(buf, op)
    elseif sprm == 68 then
      props.pic = u32(buf, op)
    elseif sprm == 71 then
      props.data = flag_on(buf, op)
    elseif sprm == 75 then
      props.ole2 = flag_on(buf, op)
    elseif sprm == 65 then
      props.deleted = flag_on(buf, op)
    elseif sprm == 80 then
      props.istd = u16(buf, op)
    elseif sprm == 93 then
      props.ftc = u16(buf, op)
    elseif sprm == 74 and len >= 3 then
      props.sym = { ftc = u16(buf, op), char = string.byte(buf, op + 2) }
    end
  end)
end

local function apply_pap6(props, buf, pos, last)
  each_sprm6(buf, pos, last, function(sprm, op, len)
    if sprm == 2 then
      props.istd = u16(buf, op)
    elseif sprm == 24 then
      props.in_table = flag_on(buf, op)
    elseif sprm == 25 then
      props.ttp = flag_on(buf, op)
    end
  end)
end

--------------------------------------------------------------------------------
-- Word 97-2003 document model
--------------------------------------------------------------------------------

-- FibRgFcLcb97 pair indexes
local FC97 = {
  Stshf = 1, SttbfFfn = 15, PlcffndRef = 2, PlcffndTxt = 3, PlcfBteChpx = 12, PlcfBtePapx = 13, Clx = 33,
  PlcSpaMom = 40, PlcfendRef = 46, PlcfendTxt = 47, DggInfo = 50, PlcftxbxTxt = 56, PlfLst = 73, PlfLfo = 74,
}

local function read_fib97(wdoc)
  local fib = {}
  local pos = 33
  local csw = u16(wdoc, pos)
  pos = pos + 2 + csw * 2
  local cslw = u16(wdoc, pos)
  local lw = pos + 2
  pos = pos + 2 + cslw * 4
  local cb_rg_fc_lcb = u16(wdoc, pos)
  local fc_base = pos + 2
  local function lw_at(i)
    return u32(wdoc, lw + i * 4)
  end
  fib.ccpText = lw_at(3)
  fib.ccpFtn = lw_at(4)
  fib.ccpHdd = lw_at(5)
  fib.ccpMcr = lw_at(6)
  fib.ccpAtn = lw_at(7)
  fib.ccpEdn = lw_at(8)
  fib.ccpTxbx = lw_at(9)
  fib.fc = function(name)
    local i = FC97[name]
    if not i or i >= cb_rg_fc_lcb then
      return 0, 0
    end
    return u32(wdoc, fc_base + i * 8), u32(wdoc, fc_base + i * 8 + 4)
  end
  return fib
end

-- Word 6/95 FIB: fixed offsets; its PLCs live in the WordDocument stream itself
local FC6 = {
  Stshf = 0x60, PlcffndRef = 0x68, PlcffndTxt = 0x70, PlcfBteChpx = 0xB8, PlcfBtePapx = 0xC0,
  SttbfFfn = 0xD0, Clx = 0x1A2,
}

local function read_fib6(wdoc)
  local fib = {
    ccpText = u32(wdoc, 0x34 + 1),
    ccpFtn = u32(wdoc, 0x38 + 1),
    ccpHdd = u32(wdoc, 0x3C + 1),
    ccpMcr = u32(wdoc, 0x40 + 1),
    ccpAtn = u32(wdoc, 0x44 + 1),
    ccpEdn = u32(wdoc, 0x48 + 1),
    fcMin = u32(wdoc, 0x18 + 1),
    complex = (math.floor(u16(wdoc, 11) / 4) % 2) == 1,
  }
  fib.fc = function(name)
    local off = FC6[name]
    if not off then
      return 0, 0
    end
    return u32(wdoc, off + 1), u32(wdoc, off + 5)
  end
  return fib
end

-- Word 6/95 text: one byte per character, decoded later with the run's font code page
local function read_text6(wdoc, fib)
  local chars, fcs = {}, {}
  local function add_range(cp0, fc, count)
    for k = 0, count - 1 do
      chars[cp0 + k + 1] = wdoc:sub(fc + k + 1, fc + k + 1)
      fcs[cp0 + k + 1] = fc + k
    end
  end
  local total = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn + fib.ccpEdn
  local fc_clx, lcb_clx = fib.fc("Clx")
  if fib.complex and lcb_clx > 0 and fc_clx + lcb_clx <= #wdoc then
    local clx = wdoc:sub(fc_clx + 1, fc_clx + lcb_clx)
    local cpos = 1
    while cpos <= #clx and string.byte(clx, cpos) == 1 do
      cpos = cpos + 3 + u16(clx, cpos + 1)
    end
    if string.byte(clx, cpos) == 2 then
      local plc = clx:sub(cpos + 5, cpos + 4 + u32(clx, cpos + 1))
      local n = math.floor((#plc - 4) / 12)
      for i = 0, n - 1 do
        local cp_start = u32(plc, 1 + i * 4)
        local cp_end = u32(plc, 1 + (i + 1) * 4)
        local fc = u32(plc, 1 + (n + 1) * 4 + i * 8 + 2) % 0x40000000
        add_range(cp_start, fc, cp_end - cp_start)
      end
      return chars, fcs
    end
  end
  add_range(0, fib.fcMin, total + 1)
  return chars, fcs
end

-- Font table: fonts[ftc] = { name, chs }
local function read_fonts(tstream, fib, word6)
  local fonts = {}
  local fc, lcb = fib.fc("SttbfFfn")
  if lcb < 4 or fc + lcb > #tstream then
    return fonts
  end
  local last = fc + lcb
  local pos = fc + 1 + (word6 and 2 or 4)
  local idx = 0
  while pos <= last do
    local cb = (string.byte(tstream, pos) or 0) + 1
    if cb < 6 or pos + cb - 1 > last then
      break
    end
    local chs = string.byte(tstream, pos + 4)
    local name
    if word6 then
      name = tstream:sub(pos + 6, pos + cb - 1):match("^[^%z]*")
    else
      local units = {}
      local q = pos + 40
      while q + 1 <= pos + cb - 1 do
        local w = u16(tstream, q)
        if w == 0 then
          break
        end
        units[#units + 1] = codepoint_to_utf8(w)
        q = q + 2
      end
      name = table.concat(units)
    end
    fonts[idx] = { name = name, chs = chs }
    idx = idx + 1
    pos = pos + cb
  end
  return fonts
end

local CHARSET_CODEPAGE = {
  [161] = 1253, [162] = 1254, [163] = 1258, [177] = 1255, [178] = 1256, [186] = 1257,
  [204] = 1251, [222] = 874, [238] = 1250,
}

-- Character stream: chars[cp+1] (UTF-8 string per character), fcs[cp+1] (byte offset in WordDocument)
local function read_text(wdoc, tstream, fib, total_cp)
  local chars, fcs = {}, {}
  local fc_clx, lcb_clx = fib.fc("Clx")
  if lcb_clx == 0 or fc_clx + lcb_clx > #tstream then
    return nil, "missing or out-of-range piece table"
  end
  local clx = tstream:sub(fc_clx + 1, fc_clx + lcb_clx)
  local cpos = 1
  while cpos <= #clx and string.byte(clx, cpos) == 1 do
    cpos = cpos + 3 + u16(clx, cpos + 1)
  end
  if string.byte(clx, cpos) ~= 2 then
    return nil, "piece table marker not found"
  end
  local lcb = u32(clx, cpos + 1)
  local plc = clx:sub(cpos + 5, cpos + 4 + lcb)
  local n = math.floor((#plc - 4) / 12)
  for i = 0, n - 1 do
    local cp_start = u32(plc, 1 + i * 4)
    local cp_end = math.min(u32(plc, 1 + (i + 1) * 4), total_cp)
    local fc_raw = u32(plc, 1 + (n + 1) * 4 + i * 8 + 2)
    local compressed = (math.floor(fc_raw / 0x40000000) % 2) == 1
    local fc = fc_raw % 0x40000000
    if compressed then
      fc = math.floor(fc / 2)
    end
    for cp = cp_start, cp_end - 1 do
      local k = cp - cp_start
      local ch
      if compressed then
        local b = string.byte(wdoc, fc + k + 1) or 32
        fcs[cp + 1] = fc + k
        if b >= 0x80 and b <= 0x9F then
          ch = codepoint_to_utf8(CP1252_MAP[b] or 32)
        else
          ch = codepoint_to_utf8(b)
        end
      else
        local w = u16(wdoc, fc + 2 * k + 1)
        fcs[cp + 1] = fc + 2 * k
        ch = codepoint_to_utf8(w)
      end
      chars[cp + 1] = ch
    end
  end
  -- Surrogate pairs arrive as two separate UTF-16 units; join them
  for i = 1, #chars - 1 do
    local a, b = chars[i], chars[i + 1]
    if a and b and #a == 3 and #b == 3 then
      local hi = (a:byte(1) % 16) * 4096 + (a:byte(2) % 64) * 64 + (a:byte(3) % 64)
      local lo = (b:byte(1) % 16) * 4096 + (b:byte(2) % 64) * 64 + (b:byte(3) % 64)
      if hi >= 0xD800 and hi <= 0xDBFF and lo >= 0xDC00 and lo <= 0xDFFF then
        chars[i] = codepoint_to_utf8(0x10000 + (hi - 0xD800) * 1024 + (lo - 0xDC00))
        chars[i + 1] = ""
      end
    end
  end
  return chars, fcs
end

-- Formatted-disk-page runs: sorted list of { fc_start, fc_end, page, offset } for CHPX or PAPX
local function read_fkp_runs(wdoc, tstream, fc, lcb, is_papx, word6)
  local runs = {}
  local pn_size = word6 and 2 or 4
  if lcb < 4 + 2 * pn_size or fc + lcb > #tstream then
    return runs
  end
  local plc = tstream:sub(fc + 1, fc + lcb)
  local n = math.floor((#plc - 4) / (4 + pn_size))
  for i = 0, n - 1 do
    local pn
    if word6 then
      pn = u16(plc, 1 + (n + 1) * 4 + i * 2)
    else
      pn = u32(plc, 1 + (n + 1) * 4 + i * 4) % 0x400000
    end
    local page = pn * 512
    if page + 512 <= #wdoc then
      local crun = string.byte(wdoc, page + 512)
      for k = 0, crun - 1 do
        local fc_start = u32(wdoc, page + 1 + k * 4)
        local fc_end = u32(wdoc, page + 1 + (k + 1) * 4)
        local entry_pos
        if is_papx then
          entry_pos = page + 1 + (crun + 1) * 4 + k * (word6 and 7 or 13)
        else
          entry_pos = page + 1 + (crun + 1) * 4 + k
        end
        local word_off = string.byte(wdoc, entry_pos) or 0
        runs[#runs + 1] = { fc_start, fc_end, page, word_off * 2 }
      end
    end
  end
  table.sort(runs, function(a, b)
    return a[1] < b[1]
  end)
  return runs
end

local function find_run(runs, fc)
  local lo, hi = 1, #runs
  while lo <= hi do
    local mid = math.floor((lo + hi) / 2)
    local r = runs[mid]
    if fc < r[1] then
      hi = mid - 1
    elseif fc >= r[2] then
      lo = mid + 1
    else
      return r
    end
  end
  return nil
end

-- Stylesheet: styles[istd] = { sti, name, base, kind, papx = {buf,pos,last}, chpx = {buf,pos,last} }
local function read_stylesheet(tstream, fib, word6)
  local styles = {}
  local fc, lcb = fib.fc("Stshf")
  if lcb < 4 or fc + lcb > #tstream then
    return styles
  end
  local st = tstream:sub(fc + 1, fc + lcb)
  local cb_stshi = u16(st, 1)
  local cstd = u16(st, 3)
  local cb_base = word6 and 10 or u16(st, 5)
  if cb_base == 0 then
    cb_base = 10
  end
  local pos = 3 + cb_stshi
  for istd = 0, cstd - 1 do
    if pos + 1 > #st then
      break
    end
    local cb_std = u16(st, pos)
    local std = pos + 2
    if cb_std > 0 and std + cb_std - 1 <= #st then
      local w0 = u16(st, std)
      local w1 = u16(st, std + 2)
      local w2 = u16(st, std + 4)
      local style = {
        sti = w0 % 4096,
        kind = w1 % 16, -- 1 paragraph, 2 character, 3 table, 4 numbering
        base = math.floor(w1 / 16),
        cupx = w2 % 16,
      }
      local p = std + cb_base
      if word6 then
        -- Pascal string of 8-bit characters plus a terminating zero
        local cch = string.byte(st, p) or 0
        style.name = st:sub(p + 1, p + cch)
        p = p + 1 + cch + 1
      else
        local cch = u16(st, p)
        local name_chars = {}
        for k = 0, cch - 1 do
          name_chars[#name_chars + 1] = codepoint_to_utf8(u16(st, p + 2 + k * 2))
        end
        style.name = table.concat(name_chars)
        p = p + 2 + cch * 2 + 2
      end
      local upx = {}
      for k = 1, style.cupx do
        if (p - std) % 2 == 1 then
          p = p + 1
        end
        local cb = u16(st, p)
        upx[k] = { st, p + 2, p + 1 + cb }
        p = p + 2 + cb
      end
      if style.kind == 1 then
        if upx[1] then
          -- paragraph UPX starts with the istd itself
          style.papx = { upx[1][1], upx[1][2] + 2, upx[1][3] }
        end
        style.chpx = upx[2]
      elseif style.kind == 2 then
        style.chpx = upx[1]
      end
      styles[istd] = style
    end
    pos = std + cb_std
  end
  return styles
end

-- Lists: lfo index (1-based) -> { [ilvl] = ordered }
local function read_lists(tstream, fib)
  local lists_by_lsid = {}
  local fc, lcb = fib.fc("PlfLst")
  if lcb >= 2 and fc + lcb <= #tstream then
    local base = fc + 1
    local c_lst = u16(tstream, base)
    local lstf_pos = base + 2
    local lvl_pos = lstf_pos + c_lst * 28
    for i = 0, c_lst - 1 do
      local lp = lstf_pos + i * 28
      local lsid = u32(tstream, lp)
      local flags = string.byte(tstream, lp + 26) or 0
      local levels = (flags % 2 == 1) and 1 or 9
      local ordered = {}
      for l = 0, levels - 1 do
        local nfc = string.byte(tstream, lvl_pos + 4)
        ordered[l] = not (nfc == 23 or nfc == 0xFF)
        local cb_chpx = string.byte(tstream, lvl_pos + 24) or 0
        local cb_papx = string.byte(tstream, lvl_pos + 25) or 0
        local xst = lvl_pos + 28 + cb_papx + cb_chpx
        local cch = u16(tstream, xst)
        lvl_pos = xst + 2 + cch * 2
      end
      lists_by_lsid[lsid] = ordered
    end
  end
  local lfos = {}
  fc, lcb = fib.fc("PlfLfo")
  if lcb >= 4 and fc + lcb <= #tstream then
    local n = u32(tstream, fc + 1)
    for i = 1, n do
      local lsid = u32(tstream, fc + 1 + 4 + (i - 1) * 16)
      lfos[i] = lists_by_lsid[lsid]
    end
  end
  return lfos
end

-- PLC of CPs with fixed-size data: returns cps (n+1 entries) and data offsets
local function read_plc(tstream, fc, lcb, cb_data)
  local cps = {}
  if lcb < 4 or fc + lcb > #tstream then
    return cps, 0
  end
  local n = math.floor((lcb - 4) / (4 + cb_data))
  for i = 0, n do
    cps[i + 1] = u32(tstream, fc + 1 + i * 4)
  end
  return cps, n, fc + 1 + (n + 1) * 4
end

-- Floating pictures (\8): spid -> BLIP via the drawing group in the table stream
local function read_floating_pictures(wdoc, tstream, fib)
  local by_cp = {}
  local fc, lcb = fib.fc("PlcSpaMom")
  local cps, n, data_pos = read_plc(tstream, fc, lcb, 26)
  if n == 0 then
    return by_cp
  end
  local dfc, dlcb = fib.fc("DggInfo")
  if dlcb == 0 or dfc + dlcb > #tstream then
    return by_cp
  end
  local first, last = dfc + 1, dfc + dlcb
  local bse = {} -- 1-based BStore entries: { pos of BSE record }
  local spid_pib = {}
  local cur_spid = nil
  local pos = first
  while pos + 7 <= last do
    local ver_inst = u16(tstream, pos)
    local rec_type = u16(tstream, pos + 2)
    local rec_len = u32(tstream, pos + 4)
    local ver = ver_inst % 16
    if ver == 0xF then
      pos = pos + 8
    else
      if rec_type == 0xF007 then
        bse[#bse + 1] = pos
      elseif rec_type == 0xF00A then
        cur_spid = u32(tstream, pos + 8)
      elseif rec_type == 0xF00B and cur_spid then
        local count = math.floor(ver_inst / 16)
        for k = 0, count - 1 do
          local opid = u16(tstream, pos + 8 + k * 6) % 16384
          if opid == 0x0104 then
            spid_pib[cur_spid] = u32(tstream, pos + 8 + k * 6 + 2)
          end
        end
      end
      pos = pos + 8 + rec_len
    end
  end
  for i = 0, n - 1 do
    local spid = u32(tstream, data_pos + i * 26)
    local pib = spid_pib[spid]
    local bpos = pib and bse[pib]
    if bpos then
      local size = u32(tstream, bpos + 8 + 20)
      local fo_delay = u32(tstream, bpos + 8 + 28)
      local data, ext
      if fo_delay ~= 0xFFFFFFFF and fo_delay < #wdoc then
        data, ext = find_blip(wdoc, fo_delay + 1, math.min(#wdoc, fo_delay + size))
      else
        data, ext = find_blip(tstream, bpos + 8 + 36, bpos + 7 + u32(tstream, bpos + 4))
      end
      if data then
        by_cp[cps[i + 1]] = { bytes = data, ext = ext, spid = spid }
      end
    end
    if not by_cp[cps[i + 1]] then
      by_cp[cps[i + 1]] = { spid = spid }
    end
  end
  return by_cp
end

-- Text-box stories: spid -> { first_cp, last_cp } in the text-box sub-document
local function read_textboxes(tstream, fib)
  local stories = {}
  local fc, lcb = fib.fc("PlcftxbxTxt")
  local cps, n, data_pos = read_plc(tstream, fc, lcb, 22)
  local base = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn + fib.ccpEdn
  for i = 0, n - 1 do
    local lid = u32(tstream, data_pos + i * 22 + 14)
    if cps[i + 2] and cps[i + 2] > cps[i + 1] then
      stories[lid] = { base + cps[i + 1], base + cps[i + 2] - 1 }
    end
  end
  return stories
end

local function itap_of(pap)
  return pap.itap or (pap.in_table and 1 or 0)
end

local function style_kind(style)
  if not style then
    return nil
  end
  if style.sti >= 1 and style.sti <= 9 then
    return "heading", style.sti
  end
  local name = (style.name or ""):lower()
  local h = name:match("^heading%s*(%d)$") or name:match("^kop%s*(%d)$") or name:match("^berschrift%s*(%d)$")
  if h then
    return "heading", tonumber(h)
  end
  if name == "title" or name == "titel" then
    return "heading", 1
  end
  if name:find("quote", 1, true) or name:find("citaat", 1, true) or name == "block text" then
    return "blockquote"
  end
  if name:find("code", 1, true) or name == "plain text" or name == "html preformatted" then
    return "code"
  end
  return nil
end

-- word6: Word 6.0/95 file (one-byte SPRMs, 8-bit text, everything in the WordDocument stream)
local function parse_doc(streams, wdoc, workspace, word6)
  local tstream, fib, chars, fcs
  local apply_chp_d = word6 and apply_chp6 or apply_chp
  local apply_pap_d = word6 and apply_pap6 or apply_pap
  if word6 then
    tstream = wdoc
    fib = read_fib6(wdoc)
    chars, fcs = read_text6(wdoc, fib)
  else
    local flags = u16(wdoc, 11)
    local use_1table = (math.floor(flags / 512) % 2) == 1
    tstream = streams[use_1table and "1Table" or "0Table"] or streams["1Table"] or streams["0Table"]
    if not tstream then
      return nil, "missing Table stream (1Table/0Table)"
    end
    fib = read_fib97(wdoc)
    local total_cp = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn + fib.ccpEdn + 1
    chars, fcs = read_text(wdoc, tstream, fib, total_cp + 1000000)
    if not chars then
      return nil, fcs
    end
  end
  local data_stream = streams["Data"] or (word6 and wdoc) or ""

  local chpx_runs = read_fkp_runs(wdoc, tstream, select(1, fib.fc("PlcfBteChpx")), select(2, fib.fc("PlcfBteChpx")), false, word6)
  local papx_runs = read_fkp_runs(wdoc, tstream, select(1, fib.fc("PlcfBtePapx")), select(2, fib.fc("PlcfBtePapx")), true, word6)
  local styles = read_stylesheet(tstream, fib, word6)
  local fonts = read_fonts(tstream, fib, word6)
  local lfos = word6 and {} or read_lists(tstream, fib)
  local floating = word6 and {} or read_floating_pictures(wdoc, tstream, fib)
  local textboxes = word6 and {} or read_textboxes(tstream, fib)
  -- Blocks from text boxes anchored in the paragraph being read; emitted right after it
  local pending_blocks = {}
  local ooxml_mod = require("nvim-docx.ooxml")
  -- Runs without a font of their own fall back to the stylesheet's default font, then the document language
  local doc_codepage = word6 and codepage_for_lid(u16(wdoc, 7)) or 1252
  local default_ftc = nil
  do
    local fc, lcb = fib.fc("Stshf")
    if lcb >= 16 and fc + 16 <= #tstream then
      default_ftc = u16(tstream, fc + 1 + 2 + 12)
    end
  end

  -- Turn one stored character into UTF-8 for the run's font: Word 6/95 bytes use the font's code
  -- page; symbol fonts (Symbol, Wingdings) map to their Unicode equivalents
  local function decode_char(ch, props)
    local font = fonts[props.ftc or default_ftc or -1]
    local fname = font and font.name or ""
    if props.sym then
      local sfont = fonts[props.sym.ftc]
      return ooxml_mod.decode_symbol(sfont and sfont.name or "Symbol", string.format("%04X", props.sym.char))
    end
    if word6 then
      local b = ch:byte()
      if font and font.chs == 2 then
        return ooxml_mod.decode_symbol(fname ~= "" and fname or "Symbol", string.format("%04X", b))
      end
      if b < 0x80 then
        return ch
      end
      return bytes_to_utf8(ch, font and CHARSET_CODEPAGE[font.chs] or doc_codepage)
    end
    if #ch == 3 then
      local cp = (ch:byte(1) % 16) * 4096 + (ch:byte(2) % 64) * 64 + (ch:byte(3) % 64)
      if cp >= 0xF020 and cp <= 0xF0FF then
        return ooxml_mod.decode_symbol(fname ~= "" and fname or "Symbol", string.format("%04X", cp))
      end
    elseif font and font.chs == 2 and #ch == 1 and ch:byte() >= 32 then
      return ooxml_mod.decode_symbol(fname, string.format("%04X", ch:byte()))
    end
    return ch
  end

  -- Style property resolution (with base-style chains, cached)
  local style_pap_cache, style_chp_cache = {}, {}
  local function style_pap(istd, depth)
    if style_pap_cache[istd] then
      return style_pap_cache[istd]
    end
    local s = styles[istd]
    local props = {}
    if s and (depth or 0) < 10 then
      if s.base ~= 4095 and s.base ~= istd then
        for k, v in pairs(style_pap(s.base, (depth or 0) + 1)) do
          props[k] = v
        end
      end
      if s.papx then
        apply_pap_d(props, s.papx[1], s.papx[2], s.papx[3])
      end
    end
    style_pap_cache[istd] = props
    return props
  end
  local function style_chp(istd, depth)
    if style_chp_cache[istd] then
      return style_chp_cache[istd]
    end
    local s = styles[istd]
    local props = {}
    if s and (depth or 0) < 10 then
      if s.base ~= 4095 and s.base ~= istd then
        for k, v in pairs(style_chp(s.base, (depth or 0) + 1)) do
          props[k] = v
        end
      end
      if s.chpx then
        apply_chp_d(props, props, s.chpx[1], s.chpx[2], s.chpx[3])
      end
    end
    style_chp_cache[istd] = props
    return props
  end

  local pap_cache = {}
  local function para_props(fc)
    local run = find_run(papx_runs, fc)
    if not run then
      return {}
    end
    local key = run[3] + run[4]
    if pap_cache[key] then
      return pap_cache[key]
    end
    local props = {}
    if run[4] > 0 then
      local p = run[3] + run[4] + 1
      local cb = string.byte(wdoc, p) or 0
      local size
      if word6 then
        size = cb * 2
        p = p + 1
      elseif cb == 0 then
        size = (string.byte(wdoc, p + 1) or 0) * 2
        p = p + 2
      else
        size = cb * 2 - 1
        p = p + 1
      end
      local istd = u16(wdoc, p)
      for k, v in pairs(style_pap(istd)) do
        props[k] = v
      end
      props.istd = istd
      apply_pap_d(props, wdoc, p + 2, p + size - 1)
    end
    pap_cache[key] = props
    return props
  end

  local chp_cache = {}
  local function char_props(fc, para_style_chp)
    local run = find_run(chpx_runs, fc)
    local key = (run and (run[3] + run[4]) or -1)
    local per_style = chp_cache[para_style_chp]
    if not per_style then
      per_style = {}
      chp_cache[para_style_chp] = per_style
    end
    if per_style[key] then
      return per_style[key]
    end
    local props = {}
    for k, v in pairs(para_style_chp) do
      props[k] = v
    end
    if run and run[4] > 0 then
      local p = run[3] + run[4] + 1
      local cb = string.byte(wdoc, p) or 0
      -- a character style (sprmCIstd) layers between paragraph style and direct formatting
      local direct = {}
      apply_chp_d(direct, para_style_chp, wdoc, p + 1, p + cb)
      if direct.istd then
        for k, v in pairs(style_chp(direct.istd)) do
          props[k] = v
        end
        direct = {}
        apply_chp_d(direct, props, wdoc, p + 1, p + cb)
      end
      for k, v in pairs(direct) do
        props[k] = v
      end
    end
    per_style[key] = props
    return props
  end

  local media_dir = workspace and (workspace .. "/media") or nil
  local image_count = 0
  local function save_picture(bytes, ext)
    image_count = image_count + 1
    local rel = string.format("media/image%d.%s", image_count, ext)
    if media_dir then
      if vim.fn.isdirectory(media_dir) == 0 then
        vim.fn.mkdir(media_dir, "p")
      end
      local wf = io.open(workspace .. "/" .. rel, "wb")
      if wf then
        wf:write(bytes)
        wf:close()
      end
    end
    return rel
  end

  local function inline_picture(pic_offset)
    if not pic_offset or pic_offset + 68 > #data_stream then
      return nil
    end
    local base = pic_offset + 1
    local lcb = u32(data_stream, base)
    local cb_header = u16(data_stream, base + 4)
    if lcb < cb_header or base + lcb - 1 > #data_stream then
      return nil
    end
    local mm = u16(data_stream, base + 6)
    local dxa_goal = u16(data_stream, base + 0x1C)
    local dya_goal = u16(data_stream, base + 0x1E)
    local mx = u16(data_stream, base + 0x20)
    local my = u16(data_stream, base + 0x22)
    local start = base + cb_header
    if mm == 0x66 then
      start = start + 1 + (string.byte(data_stream, start) or 0)
    end
    local bytes, ext = find_blip(data_stream, start, base + lcb - 1)
    if not bytes then
      return nil
    end
    local node = { type = "image", alt = "", src = save_picture(bytes, ext), _bytes = bytes }
    if dxa_goal > 0 and dya_goal > 0 then
      local w = dxa_goal * (mx > 0 and mx or 1000) / 1000 / 1440
      local h = dya_goal * (my > 0 and my or 1000) / 1000 / 1440
      node.width = string.format("%.4gin", w)
      node.height = string.format("%.4gin", h)
    end
    return node
  end

  -- Footnote / endnote reference positions
  local function note_table(ref_idx, txt_idx, sub_start)
    local fc_r, lcb_r = fib.fc(ref_idx)
    local ref_cps, n = read_plc(tstream, fc_r, lcb_r, 2)
    local fc_t, lcb_t = fib.fc(txt_idx)
    local txt_cps = read_plc(tstream, fc_t, lcb_t, 0)
    local refs = {}
    for i = 1, n do
      if txt_cps[i] and txt_cps[i + 1] then
        refs[ref_cps[i]] = { sub_start + txt_cps[i], sub_start + txt_cps[i + 1] }
      end
    end
    return refs
  end
  local ftn_refs = note_table("PlcffndRef", "PlcffndTxt", fib.ccpText)
  local edn_refs = note_table("PlcfendRef", "PlcfendTxt",
    fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn)

  local ast = { blocks = {}, footnotes = {}, metadata = {} }
  local note_counter = 0
  local parse_range -- forward

  -- Convert one paragraph's characters [cp_first, cp_last] into inlines
  local function para_inlines(cp_first, cp_last, pstyle_chp, field_stack)
    local inlines = {}
    local buf, buf_props = {}, nil
    local function flush()
      if #buf > 0 then
        -- Formatting that merely repeats the paragraph style (bold headings, ...) is the style's look,
        -- not emphasis in the text
        local function own(key)
          return (buf_props[key] and not pstyle_chp[key]) or nil
        end
        local node = {
          type = "text",
          text = table.concat(buf),
          bold = own("bold"),
          italic = own("italic"),
          strike = own("strike"),
          underline = own("underline"),
          sup = own("sup"),
          sub = own("sub"),
        }
        inlines[#inlines + 1] = node
        buf = {}
      end
    end
    local function emitting()
      for _, f in ipairs(field_stack) do
        if not f.in_result then
          return false
        end
      end
      return true
    end
    local function push_node(node)
      flush()
      inlines[#inlines + 1] = node
    end
    for cp = cp_first, cp_last do
      local ch = chars[cp + 1]
      local fc = fcs[cp + 1]
      if ch and fc then
        local props = char_props(fc, pstyle_chp)
        local b = #ch == 1 and ch:byte() or 999
        if b == 19 then
          field_stack[#field_stack + 1] = { instr = {}, in_result = false, start = nil }
        elseif b == 20 and #field_stack > 0 then
          local f = field_stack[#field_stack]
          flush()
          f.in_result = true
          f.start = #inlines + 1
        elseif b == 21 and #field_stack > 0 then
          flush()
          local f = table.remove(field_stack)
          local instr = table.concat(f.instr)
          local url = instr:match('HYPERLINK%s+"([^"]+)"') or instr:match("HYPERLINK%s+(%S+)")
          local anchor = instr:match('\\l%s+"([^"]+)"')
          if f.start and (url or anchor) then
            local inner = {}
            for k = f.start, #inlines do
              inner[#inner + 1] = inlines[k]
            end
            for k = #inlines, f.start, -1 do
              inlines[k] = nil
            end
            if #inner > 0 then
              -- the blue underline belongs to the Hyperlink style
              for _, il in ipairs(inner) do
                il.underline = nil
              end
              local target = url or ""
              if anchor then
                target = (url and url ~= "" and not url:match("^\\") and url or "") .. "#" .. anchor
              end
              inlines[#inlines + 1] = { type = "link", url = target, inlines = inner }
            end
          end
        elseif #field_stack > 0 and not field_stack[#field_stack].in_result then
          local f = field_stack[#field_stack]
          f.instr[#f.instr + 1] = ch
        elseif not emitting() then
          -- inside the instruction of an outer field
        elseif props.vanish or props.deleted then
          -- hidden text and deleted revisions are not part of the visible document
        elseif props.special and b == 1 then
          if not props.data then
            local node = inline_picture(props.pic)
            if node then
              push_node(node)
            end
          end
        elseif props.special and b == 8 then
          local shape = floating[cp]
          if shape and shape.bytes then
            push_node({ type = "image", alt = "", src = save_picture(shape.bytes, shape.ext) })
          end
          local story = shape and textboxes[shape.spid]
          if story then
            textboxes[shape.spid] = nil -- a story is shown once even if linked boxes share it
            vim.list_extend(pending_blocks, parse_range(story[1], story[2], false))
          end
        elseif props.special and b == 2 then
          local range = ftn_refs[cp] or edn_refs[cp]
          if range then
            note_counter = note_counter + 1
            local id = tostring(note_counter)
            push_node({ type = "footnote_ref", id = id })
            local note_blocks = parse_range(range[1], range[2] - 1, true)
            local note_inlines = {}
            for _, nb in ipairs(note_blocks) do
              if #note_inlines > 0 then
                note_inlines[#note_inlines + 1] = { type = "text", text = " " }
              end
              for _, il in ipairs(nb.inlines or {}) do
                note_inlines[#note_inlines + 1] = il
              end
            end
            ast.footnotes[#ast.footnotes + 1] = { id = id, inlines = markdown.trim_inlines(note_inlines, true) }
          end
        elseif props.special then
          -- other special characters (annotation refs, auto page numbers, ...) are not text
        elseif b == 9 then
          push_node({ type = "tab" })
        elseif b == 11 then
          push_node({ type = "linebreak" })
        elseif b == 30 then
          ch = "-"
        elseif b == 31 or (b < 32 and b ~= 30) then
          ch = nil
        end
        if ch and b >= 30 and not props.special and emitting() and not props.vanish and not props.deleted
          and not (#field_stack > 0 and not field_stack[#field_stack].in_result) then
          if b ~= 31 then
            if buf_props and not (buf_props.bold == props.bold and buf_props.italic == props.italic
              and buf_props.strike == props.strike and buf_props.underline == props.underline
              and buf_props.sup == props.sup and buf_props.sub == props.sub) then
              flush()
            end
            buf_props = props
            buf[#buf + 1] = decode_char(ch, props)
          end
        end
      end
    end
    flush()
    return inlines
  end

  -- Split a CP range into paragraphs and build blocks
  parse_range = function(cp_first, cp_last, is_note)
    local blocks = {}
    local field_stack = {}
    local table_rows, row_cells, cell_inlines = nil, nil, nil
    local function flush_table()
      if table_rows and #table_rows > 0 then
        local num_cols = 0
        for _, r in ipairs(table_rows) do
          num_cols = math.max(num_cols, #r)
        end
        local aligns = {}
        for c = 1, num_cols do
          aligns[c] = "left"
        end
        local rows = {}
        for r = 2, #table_rows do
          rows[#rows + 1] = table_rows[r]
        end
        blocks[#blocks + 1] = { type = "table", alignments = aligns, headers = table_rows[1], rows = rows }
      end
      table_rows, row_cells, cell_inlines = nil, nil, nil
    end

    local start = cp_first
    local cp = cp_first
    while cp <= cp_last do
      local ch = chars[cp + 1]
      local b = ch and #ch == 1 and ch:byte() or 0
      if b == 13 or b == 7 or b == 12 or b == 14 or cp == cp_last then
        local mark_fc = fcs[cp + 1] or 0
        local pap = para_props(mark_fc)
        local istd = pap.istd or 0
        local pchp = style_chp(istd)
        local last_text = (b == 13 or b == 7 or b == 12 or b == 14) and (cp - 1) or cp
        local inlines = para_inlines(start, last_text, pchp, field_stack)
        inlines = markdown.trim_inlines(inlines, itap_of(pap) >= 1 or style_kind(styles[istd]) == "heading"
          or (pap.ilfo and pap.ilfo > 0))
        local itap = itap_of(pap)
        if itap >= 1 and not is_note then
          table_rows = table_rows or {}
          row_cells = row_cells or {}
          cell_inlines = cell_inlines or {}
          local row_end = (b == 7 and itap == 1 and pap.ttp) or (itap > 1 and pap.inner_ttp)
          if itap > 1 then
            -- nested table: flatten into the outer cell
            if #inlines > 0 then
              if #cell_inlines > 0 then
                cell_inlines[#cell_inlines + 1] = { type = "text", text = row_end and "" or " " }
              end
              vim.list_extend(cell_inlines, inlines)
            end
          elseif row_end then
            table_rows[#table_rows + 1] = row_cells
            row_cells, cell_inlines = {}, {}
          else
            if #inlines > 0 then
              if #cell_inlines > 0 then
                cell_inlines[#cell_inlines + 1] = { type = "linebreak" }
              end
              vim.list_extend(cell_inlines, inlines)
            end
            if b == 7 then
              row_cells[#row_cells + 1] = cell_inlines
              cell_inlines = {}
            end
          end
        else
          flush_table()
          local kind, level = style_kind(styles[istd])
          local lfo = pap.ilfo and pap.ilfo > 0 and pap.ilfo < 0xF801 and lfos[pap.ilfo] or nil
          if #inlines > 0 then
            if kind == "heading" and not is_note then
              blocks[#blocks + 1] = { type = "heading", level = math.min(level, 6), inlines = inlines }
            elseif pap.outline and pap.outline < 9 and not is_note and not lfo then
              blocks[#blocks + 1] = { type = "heading", level = math.min(pap.outline + 1, 6), inlines = inlines }
            elseif lfo and not is_note then
              local lvl = pap.ilvl or 0
              blocks[#blocks + 1] = { type = "list_item", level = lvl, ordered = lfo[lvl] or false, inlines = inlines }
            elseif kind == "blockquote" then
              blocks[#blocks + 1] = { type = "blockquote", level = 1, inlines = inlines }
            elseif kind == "code" then
              local parts = {}
              for _, il in ipairs(inlines) do
                parts[#parts + 1] = il.type == "text" and il.text or (il.type == "tab" and "\t" or "")
              end
              local prev = blocks[#blocks]
              if prev and prev.type == "code_block" then
                prev.text = prev.text .. "\n" .. table.concat(parts)
              else
                blocks[#blocks + 1] = { type = "code_block", lang = "", text = table.concat(parts) }
              end
            else
              blocks[#blocks + 1] = { type = "paragraph", inlines = inlines }
            end
          end
        end
        if #pending_blocks > 0 and not table_rows then
          vim.list_extend(blocks, pending_blocks)
          pending_blocks = {}
        end
        start = cp + 1
      end
      cp = cp + 1
    end
    if row_cells and #row_cells > 0 then
      table_rows[#table_rows + 1] = row_cells
    end
    flush_table()
    if #pending_blocks > 0 then
      vim.list_extend(blocks, pending_blocks)
      pending_blocks = {}
    end
    return blocks
  end

  ast.blocks = parse_range(0, fib.ccpText - 1, false)

  -- Ordered list numbers continue per level until a non-list block
  local counters = {}
  for _, blk in ipairs(ast.blocks) do
    if blk.type == "list_item" then
      for lvl in pairs(counters) do
        if lvl > blk.level then
          counters[lvl] = nil
        end
      end
      if blk.ordered then
        counters[blk.level] = (counters[blk.level] or 0) + 1
        blk.number = counters[blk.level]
      end
    else
      counters = {}
    end
  end
  return ast
end

--------------------------------------------------------------------------------
-- Text-only readers for older formats (Word 6/95, Word for Windows 1/2, Write)
--------------------------------------------------------------------------------

local function text_to_ast(text)
  local blocks = {}
  local table_rows, row_cells = nil, nil
  text = text:gsub("\r\n", "\r"):gsub("\n", "\r")
  -- strip field instructions, keep results
  text = text:gsub("\19[^\19\20\21]*\20([^\19\20\21]*)\21", "%1"):gsub("\19[^\19\20\21]*\21", "")
  for para in (text .. "\r"):gmatch("([^\r]*)\r") do
    local parts = vim.split(para, "\7", { plain = true })
    if #parts > 1 then
      -- Without paragraph properties, a cell mark directly after another cell mark ends the row
      table_rows = table_rows or {}
      row_cells = row_cells or {}
      for k = 1, #parts - 1 do
        local cell = parts[k]:gsub("[%z\1-\8\11-\31]", ""):gsub("^%s+", ""):gsub("%s+$", "")
        if parts[k] == "" and k > 1 and #row_cells > 0 then
          table_rows[#table_rows + 1] = row_cells
          row_cells = {}
        else
          row_cells[#row_cells + 1] = { { type = "text", text = cell } }
        end
      end
      if parts[#parts] ~= "" then
        -- text after the last cell mark starts the next cell
        row_cells[#row_cells + 1] = { { type = "text", text = parts[#parts] } }
      end
    else
      if table_rows then
        if #row_cells > 0 then
          table_rows[#table_rows + 1] = row_cells
        end
        local aligns, cols = {}, 0
        for _, r in ipairs(table_rows) do
          cols = math.max(cols, #r)
        end
        for c = 1, cols do
          aligns[c] = "left"
        end
        local rows = {}
        for r = 2, #table_rows do
          rows[#rows + 1] = table_rows[r]
        end
        blocks[#blocks + 1] = { type = "table", alignments = aligns, headers = table_rows[1] or {}, rows = rows }
        table_rows, row_cells = nil, nil
      end
      local clean = para:gsub("\11", "\n"):gsub("[%z\1-\8\12-\31]", ""):gsub("\9", "\t")
      if clean:match("%S") then
        local inlines = {}
        for k, line in ipairs(vim.split(clean, "\n", { plain = true })) do
          if k > 1 then
            inlines[#inlines + 1] = { type = "linebreak" }
          end
          for t, piece in ipairs(vim.split(line, "\t", { plain = true })) do
            if t > 1 then
              inlines[#inlines + 1] = { type = "tab" }
            end
            inlines[#inlines + 1] = { type = "text", text = piece }
          end
        end
        blocks[#blocks + 1] = { type = "paragraph", inlines = markdown.trim_inlines(inlines, false) }
      end
    end
  end
  return { blocks = blocks, footnotes = {}, metadata = {} }
end

-- Word for Windows 1.x/2.0 (no OLE container): text runs from fcMin
local function parse_word2(data)
  local fc_min = u32(data, 0x18 + 1)
  local ccp_text = u32(data, 0x34 + 1)
  if fc_min <= 0 or fc_min >= #data then
    fc_min = 0x180
  end
  if ccp_text <= 0 or fc_min + ccp_text > #data then
    ccp_text = u32(data, 0x1C + 1) - fc_min
  end
  return text_to_ast(cp1252_to_utf8(data:sub(fc_min + 1, fc_min + math.max(ccp_text, 0))))
end

-- Windows Write (.wri): text from offset 128 to fcMac
local function parse_write(data)
  local fc_mac = u32(data, 14 + 1)
  if fc_mac <= 128 or fc_mac > #data then
    fc_mac = #data
  end
  return text_to_ast(cp1252_to_utf8(data:sub(129, fc_mac)))
end

function M.doc_to_ast(filepath, workspace)
  local f, err = io.open(filepath, "rb")
  if not f then
    return nil, "Cannot open .doc file: " .. tostring(err)
  end
  local data = f:read("*a")
  f:close()

  local kind = M.detect_legacy(filepath)
  if kind == "word2" then
    return parse_word2(data)
  elseif kind == "write" then
    return parse_write(data)
  end

  local streams, stream_err = read_ole2_streams(data)
  if not streams then
    return nil, stream_err
  end
  if streams["EncryptedPackage"] or streams["EncryptionInfo"] then
    return nil, "Document is password-protected (encrypted); remove the password in Word first"
  end

  local wdoc = streams["WordDocument"]
  if not wdoc or #wdoc < 64 then
    return nil, "Not a Word document: the OLE container has no WordDocument stream"
  end
  local ident = u16(wdoc, 1)
  if ident ~= 0xA5EC and ident ~= 0xA5DC then
    return nil, string.format("Unsupported Word binary signature: 0x%04X", ident)
  end
  local flags = u16(wdoc, 11)
  if (math.floor(flags / 256) % 2) == 1 then
    return nil, "Document is password-protected (encrypted); remove the password in Word first"
  end

  local nfib = u16(wdoc, 3)
  if nfib < 101 then
    return nil, string.format("Unsupported Word version (nFib %d)", nfib)
  end
  local ok, ast, perr = pcall(parse_doc, streams, wdoc, workspace, nfib <= 105)
  if not ok then
    return nil, "Failed to read Word document: " .. tostring(ast)
  end
  if not ast then
    return nil, perr
  end
  return ast
end

M._read_streams = read_ole2_streams

return M
