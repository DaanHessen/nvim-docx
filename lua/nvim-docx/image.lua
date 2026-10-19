-- Pure-Lua image header parser, EMU converter, and MIME type mapper for OOXML drawings.

local M = {}

local EMU_PER_INCH = 914400
local EMU_PER_CM = 360000
local EMU_PER_MM = 36000
local EMU_PER_PT = 12700
local EMU_PER_PC = 152400
local EMU_PER_PX = 9525 -- 96 DPI
local MAX_WIDTH_EMU = 5486400 -- 6.0 inches max default inline width

M.EMU_PER_INCH = EMU_PER_INCH
M.EMU_PER_PX = EMU_PER_PX

local CONTENT_TYPES = {
  png = "image/png",
  jpg = "image/jpeg",
  jpeg = "image/jpeg",
  gif = "image/gif",
  bmp = "image/bmp",
  svg = "image/svg+xml",
  webp = "image/webp",
  tif = "image/tiff",
  tiff = "image/tiff",
  emf = "image/x-emf",
  wmf = "image/x-wmf",
}

function M.content_type_for_ext(ext)
  if not ext then
    return "application/octet-stream"
  end
  ext = ext:lower():gsub("^%.", "")
  return CONTENT_TYPES[ext] or ("image/" .. ext)
end

local function read_u16_be(s, pos)
  local b1, b2 = string.byte(s, pos, pos + 1)
  if not b2 then return nil end
  return b1 * 256 + b2
end

local function read_u32_be(s, pos)
  local b1, b2, b3, b4 = string.byte(s, pos, pos + 3)
  if not b4 then return nil end
  return b1 * 16777216 + b2 * 65536 + b3 * 256 + b4
end

local function read_u16_le(s, pos)
  local b1, b2 = string.byte(s, pos, pos + 1)
  if not b2 then return nil end
  return b1 + b2 * 256
end

local function read_u32_le(s, pos)
  local b1, b2, b3, b4 = string.byte(s, pos, pos + 3)
  if not b4 then return nil end
  return b1 + b2 * 256 + b3 * 65536 + b4 * 16777216
end

-- Inspect image bytes to extract pixel dimensions (width, height)
function M.get_pixel_dimensions(bytes)
  if not bytes or #bytes < 10 then
    return nil, nil
  end

  -- 1. PNG: 89 50 4E 47 0D 0A 1A 0A + IHDR
  if bytes:sub(1, 8) == "\137PNG\r\n\26\n" and #bytes >= 24 then
    local w = read_u32_be(bytes, 17)
    local h = read_u32_be(bytes, 21)
    if w and h and w > 0 and h > 0 then
      return w, h
    end
  end

  -- 2. GIF: GIF87a or GIF89a
  local hdr6 = bytes:sub(1, 6)
  if (hdr6 == "GIF87a" or hdr6 == "GIF89a") and #bytes >= 10 then
    local w = read_u16_le(bytes, 7)
    local h = read_u16_le(bytes, 9)
    if w and h and w > 0 and h > 0 then
      return w, h
    end
  end

  -- 3. BMP: "BM"
  if bytes:sub(1, 2) == "BM" and #bytes >= 26 then
    local w = read_u32_le(bytes, 19)
    local h = read_u32_le(bytes, 23)
    if w and h then
      if h >= 0x80000000 then
        h = 0x100000000 - h
      end
      if w > 0 and h > 0 then
        return w, h
      end
    end
  end

  -- 4. JPEG: FF D8
  if string.byte(bytes, 1) == 0xFF and string.byte(bytes, 2) == 0xD8 then
    local pos = 3
    local len = #bytes
    while pos + 3 <= len do
      if string.byte(bytes, pos) ~= 0xFF then
        pos = pos + 1
      else
        while pos <= len and string.byte(bytes, pos) == 0xFF do
          pos = pos + 1
        end
        if pos > len then break end
        local marker = string.byte(bytes, pos)
        pos = pos + 1
        if marker == 0xD8 or marker == 0xD9 or (marker >= 0xD0 and marker <= 0xD7) or marker == 0x01 then
          -- Standalone marker without length
        else
          local seg_len = read_u16_be(bytes, pos)
          if not seg_len or seg_len < 2 then break end
          if marker >= 0xC0 and marker <= 0xCF and marker ~= 0xC4 and marker ~= 0xC8 and marker ~= 0xCC then
            if pos + 6 <= len then
              local h = read_u16_be(bytes, pos + 3)
              local w = read_u16_be(bytes, pos + 5)
              if w and h and w > 0 and h > 0 then
                return w, h
              end
            end
            break
          end
          pos = pos + seg_len
        end
      end
    end
  end

  -- 5. WebP: RIFF....WEBP
  if bytes:sub(1, 4) == "RIFF" and bytes:sub(9, 12) == "WEBP" and #bytes >= 30 then
    local chunk = bytes:sub(13, 16)
    if chunk == "VP8 " then
      local w = read_u16_le(bytes, 27)
      local h = read_u16_le(bytes, 29)
      if w and h then
        return (w % 16384), (h % 16384)
      end
    elseif chunk == "VP8X" then
      local b1, b2, b3 = string.byte(bytes, 25, 27)
      local h1, h2, h3 = string.byte(bytes, 28, 30)
      if b3 and h3 then
        local w = 1 + b1 + b2 * 256 + b3 * 65536
        local h = 1 + h1 + h2 * 256 + h3 * 65536
        return w, h
      end
    end
  end

  return nil, nil
end

function M.parse_length_to_emu(str)
  if not str or str == "" then
    return nil
  end
  str = tostring(str):gsub("^%s+", ""):gsub("%s+$", "")
  local num_str, unit = str:match("^([%d%.]+)%s*(%a*)$")
  local val = tonumber(num_str)
  if not val or val <= 0 then
    return nil
  end
  unit = (unit or ""):lower()
  if unit == "in" or unit == "inch" or unit == "inches" then
    return math.floor(val * EMU_PER_INCH + 0.5)
  elseif unit == "cm" then
    return math.floor(val * EMU_PER_CM + 0.5)
  elseif unit == "mm" then
    return math.floor(val * EMU_PER_MM + 0.5)
  elseif unit == "pt" then
    return math.floor(val * EMU_PER_PT + 0.5)
  elseif unit == "pc" then
    return math.floor(val * EMU_PER_PC + 0.5)
  elseif unit == "px" or unit == "" then
    return math.floor(val * EMU_PER_PX + 0.5)
  elseif unit == "emu" then
    return math.floor(val + 0.5)
  end
  return nil
end

function M.emu_to_inches_str(emu)
  local num = tonumber(emu)
  if not num or num <= 0 then
    return nil
  end
  local inches = num / EMU_PER_INCH
  local formatted = string.format("%.4f", inches):gsub("0+$", ""):gsub("%.$", "")
  return formatted .. "in"
end

function M.resolve_dimensions_emu(image_bytes, width_str, height_str)
  local w_emu = M.parse_length_to_emu(width_str)
  local h_emu = M.parse_length_to_emu(height_str)
  if w_emu and h_emu then
    return w_emu, h_emu
  end

  local px_w, px_h = M.get_pixel_dimensions(image_bytes)
  if w_emu and not h_emu then
    if px_w and px_h and px_w > 0 then
      return w_emu, math.floor(w_emu * (px_h / px_w) + 0.5)
    end
    return w_emu, w_emu
  elseif h_emu and not w_emu then
    if px_w and px_h and px_h > 0 then
      return math.floor(h_emu * (px_w / px_h) + 0.5), h_emu
    end
    return h_emu, h_emu
  end

  if px_w and px_h and px_w > 0 and px_h > 0 then
    local cw = px_w * EMU_PER_PX
    local ch = px_h * EMU_PER_PX
    if cw > MAX_WIDTH_EMU then
      local scale = MAX_WIDTH_EMU / cw
      cw = MAX_WIDTH_EMU
      ch = math.floor(ch * scale + 0.5)
    end
    return cw, ch
  end

  -- Default fallback: 3.0in x 2.0in
  return 2743200, 1828800
end

return M
