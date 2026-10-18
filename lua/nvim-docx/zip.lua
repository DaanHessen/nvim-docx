-- Pure-Lua PKZIP reader/writer with RFC 1951 DEFLATE/INFLATE and IEEE 802.3 CRC-32.
-- Zero external binary or C library dependencies.

local M = {}

local ok_bit, bit_lib = pcall(require, "bit")

local band, bor, bxor, lshift, rshift
if ok_bit and bit_lib then
  band = bit_lib.band
  bor = bit_lib.bor
  bxor = bit_lib.bxor
  lshift = bit_lib.lshift
  rshift = bit_lib.rshift
else
  -- Pure-Lua fallback for bitwise operations
  band = function(a, b)
    local res, p = 0, 1
    a, b = a % 4294967296, b % 4294967296
    while a > 0 and b > 0 do
      local ra, rb = a % 2, b % 2
      if ra == 1 and rb == 1 then res = res + p end
      a, b, p = (a - ra) / 2, (b - rb) / 2, p * 2
    end
    return res
  end
  bor = function(a, b)
    local res, p = 0, 1
    a, b = a % 4294967296, b % 4294967296
    while a > 0 or b > 0 do
      local ra, rb = a % 2, b % 2
      if ra == 1 or rb == 1 then res = res + p end
      a, b, p = (a - ra) / 2, (b - rb) / 2, p * 2
    end
    return res
  end
  bxor = function(a, b)
    local res, p = 0, 1
    a, b = a % 4294967296, b % 4294967296
    while a > 0 or b > 0 do
      local ra, rb = a % 2, b % 2
      if ra ~= rb then res = res + p end
      a, b, p = (a - ra) / 2, (b - rb) / 2, p * 2
    end
    return res
  end
  lshift = function(a, n)
    return (a * (2 ^ n)) % 4294967296
  end
  rshift = function(a, n)
    return math.floor((a % 4294967296) / (2 ^ n))
  end
end

local function to_u32(n)
  if n < 0 then
    return n + 4294967296
  end
  return n % 4294967296
end

-- Powers of 2 lookup table (0..32)
local POW2 = {}
for i = 0, 32 do
  POW2[i] = 2 ^ i
end

-- Precompute IEEE 802.3 CRC-32 table (polynomial 0xEDB88320)
local CRC_TABLE = {}
for i = 0, 255 do
  local c = i
  for _ = 1, 8 do
    if band(c, 1) ~= 0 then
      c = bxor(rshift(c, 1), 0xEDB88320)
    else
      c = rshift(c, 1)
    end
  end
  CRC_TABLE[i] = to_u32(c)
end

function M.crc32(data)
  local crc = 0xFFFFFFFF
  local len = #data
  for i = 1, len do
    local b = string.byte(data, i)
    local idx = band(bxor(crc, b), 0xFF)
    crc = bxor(rshift(crc, 8), CRC_TABLE[idx])
  end
  return to_u32(bxor(crc, 0xFFFFFFFF))
end

-- Little-endian integer readers/writers
local function read_u16(s, pos)
  local b1, b2 = string.byte(s, pos, pos + 1)
  if not b2 then return nil end
  return b1 + b2 * 256
end

local function read_u32(s, pos)
  local b1, b2, b3, b4 = string.byte(s, pos, pos + 3)
  if not b4 then return nil end
  return b1 + b2 * 256 + b3 * 65536 + b4 * 16777216
end

local function write_u16(n)
  n = math.floor(n or 0) % 65536
  local b1 = n % 256
  local b2 = math.floor(n / 256) % 256
  return string.char(b1, b2)
end

local function write_u32(n)
  n = to_u32(math.floor(n or 0))
  local b1 = n % 256
  local b2 = math.floor(n / 256) % 256
  local b3 = math.floor(n / 65536) % 256
  local b4 = math.floor(n / 16777216) % 256
  return string.char(b1, b2, b3, b4)
end

--------------------------------------------------------------------------------
-- RFC 1951 INFLATE (Decompression)
--------------------------------------------------------------------------------

local LEN_BASE = {
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
  35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258,
}
local LEN_EXTRA = {
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
  3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
}
local DIST_BASE = {
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
  257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
  8193, 12289, 16385, 24577,
}
local DIST_EXTRA = {
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
  7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
}
local CLEN_ORDER = {
  16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
}

local function build_huffman_tree(lengths, num_syms)
  local max_bits = 0
  local bl_count = {}
  for i = 0, 15 do bl_count[i] = 0 end

  for i = 1, num_syms do
    local l = lengths[i] or 0
    if l > 0 then
      bl_count[l] = bl_count[l] + 1
      if l > max_bits then max_bits = l end
    end
  end

  if max_bits == 0 then
    return { max_bits = 0, lookup = {} }
  end

  local next_code = {}
  local code = 0
  bl_count[0] = 0
  for bits = 1, 15 do
    code = (code + bl_count[bits - 1]) * 2
    next_code[bits] = code
  end

  -- Map (bits, canonical_code) -> symbol (0-indexed)
  -- For fast LSB-first lookup, store at key = (1 << bits) + reversed_code
  local lookup = {}
  for sym = 0, num_syms - 1 do
    local l = lengths[sym + 1] or 0
    if l > 0 then
      local c = next_code[l]
      next_code[l] = c + 1
      -- Reverse `l` bits of `c`
      local rev = 0
      local tmp = c
      for _ = 1, l do
        rev = rev * 2 + (tmp % 2)
        tmp = math.floor(tmp / 2)
      end
      lookup[POW2[l] + rev] = sym
    end
  end

  return { max_bits = max_bits, lookup = lookup }
end

-- Prebuild fixed Huffman trees for btype == 1
local FIXED_LIT_TREE, FIXED_DIST_TREE
do
  local lit_lengths = {}
  for i = 0, 143 do lit_lengths[i + 1] = 8 end
  for i = 144, 255 do lit_lengths[i + 1] = 9 end
  for i = 256, 279 do lit_lengths[i + 1] = 7 end
  for i = 280, 287 do lit_lengths[i + 1] = 8 end
  FIXED_LIT_TREE = build_huffman_tree(lit_lengths, 288)

  local dist_lengths = {}
  for i = 0, 31 do dist_lengths[i + 1] = 5 end
  FIXED_DIST_TREE = build_huffman_tree(dist_lengths, 32)
end

function M.inflate(data)
  local data_len = #data
  local pos = 1
  local bit_buf = 0
  local bit_cnt = 0

  local function read_bits(n)
    if n == 0 then return 0 end
    while bit_cnt < n do
      if pos > data_len then
        error("inflate: unexpected end of stream")
      end
      local b = string.byte(data, pos)
      pos = pos + 1
      bit_buf = bit_buf + b * POW2[bit_cnt]
      bit_cnt = bit_cnt + 8
    end
    local p = POW2[n]
    local val = bit_buf % p
    bit_buf = math.floor(bit_buf / p)
    bit_cnt = bit_cnt - n
    return val
  end

  local function decode_symbol(tree)
    local lookup = tree.lookup
    local max_bits = tree.max_bits
    if not max_bits or max_bits == 0 then
      error("inflate: attempt to decode symbol from empty huffman tree")
    end
    while bit_cnt < max_bits and pos <= data_len do
      local b = string.byte(data, pos)
      pos = pos + 1
      bit_buf = bit_buf + b * POW2[bit_cnt]
      bit_cnt = bit_cnt + 8
    end
    local code = 0
    local p = 1
    for len = 1, max_bits do
      if bit_cnt == 0 then
        error("inflate: unexpected end of stream in huffman code")
      end
      local bit_val = bit_buf % 2
      bit_buf = (bit_buf - bit_val) / 2
      bit_cnt = bit_cnt - 1
      code = code + bit_val * p
      p = p * 2
      local sym = lookup[p + code]
      if sym ~= nil then
        return sym
      end
    end
    error("inflate: invalid huffman code")
  end

  local out = {}
  local out_len = 0

  repeat
    local bfinal = read_bits(1)
    local btype = read_bits(2)

    if btype == 0 then
      -- Uncompressed block: align to byte boundary
      bit_buf = 0
      bit_cnt = 0
      if pos + 3 > data_len then
        error("inflate: truncated uncompressed block header")
      end
      local len = read_u16(data, pos)
      local nlen = read_u16(data, pos + 2)
      pos = pos + 4
      if bxor(len, nlen) ~= 0xFFFF then
        error("inflate: invalid uncompressed block length complement")
      end
      if pos + len - 1 > data_len then
        error("inflate: truncated uncompressed block data")
      end
      for i = pos, pos + len - 1 do
        out_len = out_len + 1
        out[out_len] = string.byte(data, i)
      end
      pos = pos + len
    elseif btype == 1 or btype == 2 then
      local lit_tree, dist_tree
      if btype == 1 then
        lit_tree = FIXED_LIT_TREE
        dist_tree = FIXED_DIST_TREE
      else
        local hlit = read_bits(5) + 257
        local hdist = read_bits(5) + 1
        local hclen = read_bits(4) + 4
        local clen_lengths = {}
        for i = 1, 19 do clen_lengths[i] = 0 end
        for i = 1, hclen do
          clen_lengths[CLEN_ORDER[i] + 1] = read_bits(3)
        end
        local clen_tree = build_huffman_tree(clen_lengths, 19)

        local all_lengths = {}
        local total_syms = hlit + hdist
        local idx = 1
        while idx <= total_syms do
          local sym = decode_symbol(clen_tree)
          if sym < 16 then
            all_lengths[idx] = sym
            idx = idx + 1
          elseif sym == 16 then
            local prev = all_lengths[idx - 1] or 0
            local rep = 3 + read_bits(2)
            for _ = 1, rep do
              all_lengths[idx] = prev
              idx = idx + 1
            end
          elseif sym == 17 then
            local rep = 3 + read_bits(3)
            for _ = 1, rep do
              all_lengths[idx] = 0
              idx = idx + 1
            end
          elseif sym == 18 then
            local rep = 11 + read_bits(7)
            for _ = 1, rep do
              all_lengths[idx] = 0
              idx = idx + 1
            end
          else
            error("inflate: invalid code length symbol " .. tostring(sym))
          end
        end

        local lit_lengths = {}
        for i = 1, hlit do lit_lengths[i] = all_lengths[i] end
        local dist_lengths = {}
        for i = 1, hdist do dist_lengths[i] = all_lengths[hlit + i] end

        lit_tree = build_huffman_tree(lit_lengths, hlit)
        dist_tree = build_huffman_tree(dist_lengths, hdist)
      end

      while true do
        local sym = decode_symbol(lit_tree)
        if sym < 256 then
          out_len = out_len + 1
          out[out_len] = sym
        elseif sym == 256 then
          break
        else
          local len_idx = sym - 256
          local length = LEN_BASE[len_idx] + read_bits(LEN_EXTRA[len_idx])
          local dist_sym = decode_symbol(dist_tree)
          local dist = DIST_BASE[dist_sym + 1] + read_bits(DIST_EXTRA[dist_sym + 1])
          local src = out_len - dist + 1
          if src < 1 then
            error("inflate: distance back-reference out of bounds")
          end
          for _ = 1, length do
            out_len = out_len + 1
            out[out_len] = out[src]
            src = src + 1
          end
        end
      end
    else
      error("inflate: invalid block type 3")
    end
  until bfinal == 1

  -- Assemble byte array into string in chunks
  local chunks = {}
  local chunk_size = 4096
  local unpack_fn = table.unpack or unpack
  for i = 1, out_len, chunk_size do
    local j = math.min(i + chunk_size - 1, out_len)
    chunks[#chunks + 1] = string.char(unpack_fn(out, i, j))
  end
  return table.concat(chunks)
end

--------------------------------------------------------------------------------
-- RFC 1951 DEFLATE (Compression using Fixed Huffman + LZ77)
--------------------------------------------------------------------------------

local FIXED_LIT_CODE = {}
local FIXED_LIT_LEN = {}
local FIXED_DIST_CODE = {}
local LENGTH_TO_SYM = {}
local LENGTH_TO_EXTRA_VAL = {}
local LENGTH_TO_EXTRA_BITS = {}

do
  local function rev_bits(val, n)
    local r = 0
    for _ = 1, n do
      r = r * 2 + (val % 2)
      val = math.floor(val / 2)
    end
    return r
  end

  for i = 0, 143 do
    FIXED_LIT_CODE[i] = rev_bits(0x30 + i, 8)
    FIXED_LIT_LEN[i] = 8
  end
  for i = 144, 255 do
    FIXED_LIT_CODE[i] = rev_bits(0x190 + (i - 144), 9)
    FIXED_LIT_LEN[i] = 9
  end
  for i = 256, 279 do
    FIXED_LIT_CODE[i] = rev_bits(i - 256, 7)
    FIXED_LIT_LEN[i] = 7
  end
  for i = 280, 287 do
    FIXED_LIT_CODE[i] = rev_bits(0xC0 + (i - 280), 8)
    FIXED_LIT_LEN[i] = 8
  end

  for i = 0, 29 do
    FIXED_DIST_CODE[i] = rev_bits(i, 5)
  end

  for sym_idx = 1, #LEN_BASE do
    local base = LEN_BASE[sym_idx]
    local extra = LEN_EXTRA[sym_idx]
    local count = POW2[extra]
    for offset = 0, count - 1 do
      local l = base + offset
      if l <= 258 and not LENGTH_TO_SYM[l] then
        LENGTH_TO_SYM[l] = 256 + sym_idx
        LENGTH_TO_EXTRA_VAL[l] = offset
        LENGTH_TO_EXTRA_BITS[l] = extra
      end
    end
  end
  LENGTH_TO_SYM[258] = 285
  LENGTH_TO_EXTRA_VAL[258] = 0
  LENGTH_TO_EXTRA_BITS[258] = 0
end

local function find_dist_code(dist)
  -- Binary search in DIST_BASE (1..30)
  local lo, hi = 1, 30
  while lo < hi do
    local mid = math.floor((lo + hi + 1) / 2)
    if DIST_BASE[mid] <= dist then
      lo = mid
    else
      hi = mid - 1
    end
  end
  local sym = lo - 1
  return sym, dist - DIST_BASE[lo], DIST_EXTRA[lo]
end

function M.deflate(data)
  local n = #data
  local out = {}
  local out_len = 0
  local bit_buf = 0
  local bit_cnt = 0

  local function write_bits(val, nbits)
    if nbits == 0 then return end
    bit_buf = bit_buf + val * POW2[bit_cnt]
    bit_cnt = bit_cnt + nbits
    while bit_cnt >= 8 do
      out_len = out_len + 1
      out[out_len] = bit_buf % 256
      bit_buf = math.floor(bit_buf / 256)
      bit_cnt = bit_cnt - 8
    end
  end

  -- Write header for final Fixed Huffman block (bfinal = 1, btype = 1)
  write_bits(1, 1)
  write_bits(1, 2)

  local head = {}
  local prev = {}
  local max_chain = 16
  local window_size = 32768

  local pos = 1
  while pos <= n do
    local best_len = 0
    local best_dist = 0

    if pos + 2 <= n then
      local b1, b2, b3 = string.byte(data, pos, pos + 2)
      local h = (b1 * 65599 + b2 * 257 + b3) % 8192
      local cand = head[h]
      prev[pos] = cand
      head[h] = pos

      local chain_len = 0
      local max_len = math.min(258, n - pos + 1)
      while cand and (pos - cand) <= window_size and chain_len < max_chain do
        chain_len = chain_len + 1
        if string.byte(data, cand + best_len) == string.byte(data, pos + best_len) then
          local l = 0
          while l < max_len and string.byte(data, cand + l) == string.byte(data, pos + l) do
            l = l + 1
          end
          if l >= 3 and l > best_len then
            best_len = l
            best_dist = pos - cand
            if l == max_len or l >= 64 then
              break
            end
          end
        end
        cand = prev[cand]
      end
    end

    if best_len >= 3 then
      local lsym = LENGTH_TO_SYM[best_len]
      write_bits(FIXED_LIT_CODE[lsym], FIXED_LIT_LEN[lsym])
      local lextra_bits = LENGTH_TO_EXTRA_BITS[best_len]
      if lextra_bits > 0 then
        write_bits(LENGTH_TO_EXTRA_VAL[best_len], lextra_bits)
      end

      local dsym, dextra_val, dextra_bits = find_dist_code(best_dist)
      write_bits(FIXED_DIST_CODE[dsym], 5)
      if dextra_bits > 0 then
        write_bits(dextra_val, dextra_bits)
      end

      -- Update hash table for intermediate bytes of short matches
      if best_len <= 32 then
        for step = 1, best_len - 1 do
          local p = pos + step
          if p + 2 <= n then
            local b1, b2, b3 = string.byte(data, p, p + 2)
            local h = (b1 * 65599 + b2 * 257 + b3) % 8192
            prev[p] = head[h]
            head[h] = p
          end
        end
      end
      pos = pos + best_len
    else
      local b = string.byte(data, pos)
      write_bits(FIXED_LIT_CODE[b], FIXED_LIT_LEN[b])
      pos = pos + 1
    end
  end

  -- End of block symbol (256)
  write_bits(FIXED_LIT_CODE[256], FIXED_LIT_LEN[256])

  -- Flush remaining bits
  if bit_cnt > 0 then
    out_len = out_len + 1
    out[out_len] = bit_buf % 256
  end

  local chunks = {}
  local chunk_size = 4096
  local unpack_fn = table.unpack or unpack
  for i = 1, out_len, chunk_size do
    local j = math.min(i + chunk_size - 1, out_len)
    chunks[#chunks + 1] = string.char(unpack_fn(out, i, j))
  end
  return table.concat(chunks)
end

--------------------------------------------------------------------------------
-- PKZIP Reader & Writer
--------------------------------------------------------------------------------

local SIG_LOCAL_FILE = 0x04034B50
local SIG_CENTRAL_DIR = 0x02014B50
local SIG_EOCD = 0x06054B50

function M.read_zip_bytes(bytes)
  if not bytes or #bytes < 22 then
    return nil, "Invalid or truncated ZIP archive (too small)"
  end

  -- Check local file signature at start (or empty zip EOCD)
  local first_sig = read_u32(bytes, 1)
  if first_sig ~= SIG_LOCAL_FILE and first_sig ~= SIG_EOCD then
    return nil, "Not a valid ZIP/DOCX archive (invalid magic bytes)"
  end

  -- Locate End of Central Directory (EOCD) record scanning backward from end
  local len = #bytes
  local min_eocd_pos = math.max(1, len - 65557)
  local eocd_pos = nil
  for pos = len - 21, min_eocd_pos, -1 do
    if read_u32(bytes, pos) == SIG_EOCD then
      local comment_len = read_u16(bytes, pos + 20)
      if comment_len and (pos + 21 + comment_len == len) then
        eocd_pos = pos
        break
      end
      if not eocd_pos then
        eocd_pos = pos
      end
    end
  end

  if not eocd_pos then
    return nil, "Invalid ZIP archive: End of Central Directory record not found"
  end

  local total_entries = read_u16(bytes, eocd_pos + 10)
  local cd_size = read_u32(bytes, eocd_pos + 12)
  local cd_offset = read_u32(bytes, eocd_pos + 16)

  if not total_entries or not cd_size or not cd_offset or (cd_offset + cd_size > len + 1) then
    return nil, "Invalid ZIP archive: Central Directory out of bounds"
  end

  local entries = {}
  local files = {}
  local order = {}
  local pos = cd_offset + 1

  for _ = 1, total_entries do
    if pos + 45 > len then
      return nil, "Invalid ZIP archive: truncated Central Directory header"
    end
    local sig = read_u32(bytes, pos)
    if sig ~= SIG_CENTRAL_DIR then
      return nil, "Invalid ZIP archive: bad Central Directory signature"
    end

    local flags = read_u16(bytes, pos + 8)
    local method = read_u16(bytes, pos + 10)
    local dos_time = read_u16(bytes, pos + 12)
    local dos_date = read_u16(bytes, pos + 14)
    local crc = read_u32(bytes, pos + 16)
    local comp_size = read_u32(bytes, pos + 20)
    local uncomp_size = read_u32(bytes, pos + 24)
    local name_len = read_u16(bytes, pos + 28)
    local extra_len = read_u16(bytes, pos + 30)
    local comment_len = read_u16(bytes, pos + 32)
    local local_offset = read_u32(bytes, pos + 42)

    local name_start = pos + 46
    local name_end = name_start + name_len - 1
    if name_end > len then
      return nil, "Invalid ZIP archive: truncated entry filename"
    end
    local raw_name = string.sub(bytes, name_start, name_end)
    -- Normalize backslashes from Windows-created archives to forward slashes
    local name = raw_name:gsub("\\", "/")
    local extra_start = name_end + 1
    local extra_end = extra_start + extra_len - 1
    pos = extra_end + 1 + comment_len

    -- Parse ZIP64 Extended Information Extra Field (0x0001) if 32-bit fields are 0xFFFFFFFF
    if (uncomp_size == 0xFFFFFFFF or comp_size == 0xFFFFFFFF or local_offset == 0xFFFFFFFF) and extra_len >= 4 then
      local epos = extra_start
      while epos + 3 <= extra_end and epos + 3 <= len do
        local header_id = read_u16(bytes, epos)
        local data_size = read_u16(bytes, epos + 2)
        epos = epos + 4
        if not header_id or not data_size or epos + data_size - 1 > extra_end then
          break
        end
        if header_id == 0x0001 then
          local zpos = epos
          if uncomp_size == 0xFFFFFFFF and zpos + 7 <= epos + data_size - 1 then
            local lo = read_u32(bytes, zpos)
            local hi = read_u32(bytes, zpos + 4)
            uncomp_size = lo + hi * 4294967296
            zpos = zpos + 8
          end
          if comp_size == 0xFFFFFFFF and zpos + 7 <= epos + data_size - 1 then
            local lo = read_u32(bytes, zpos)
            local hi = read_u32(bytes, zpos + 4)
            comp_size = lo + hi * 4294967296
            zpos = zpos + 8
          end
          if local_offset == 0xFFFFFFFF and zpos + 7 <= epos + data_size - 1 then
            local lo = read_u32(bytes, zpos)
            local hi = read_u32(bytes, zpos + 4)
            local_offset = lo + hi * 4294967296
          end
          break
        end
        epos = epos + data_size
      end
    end

    -- Locate data in Local File Header
    local lpos = local_offset + 1
    if lpos + 29 > len or read_u32(bytes, lpos) ~= SIG_LOCAL_FILE then
      return nil, "Invalid ZIP archive: bad Local File Header for " .. tostring(name)
    end
    local lname_len = read_u16(bytes, lpos + 26)
    local lextra_len = read_u16(bytes, lpos + 28)
    local data_start = lpos + 30 + lname_len + lextra_len
    local data_end = data_start + comp_size - 1

    if data_end > len then
      return nil, "Invalid ZIP archive: truncated file data for " .. tostring(name)
    end

    local raw_data = string.sub(bytes, data_start, data_end)
    local content = ""
    if name:sub(-1) ~= "/" then
      if method == 0 then
        content = raw_data
      elseif method == 8 then
        local ok, inflated = pcall(M.inflate, raw_data)
        if not ok then
          return nil, "Failed to decompress '" .. tostring(name) .. "': " .. tostring(inflated)
        end
        content = inflated
      else
        return nil, string.format("Unsupported ZIP compression method %d for '%s'", method, name)
      end

      if #content ~= uncomp_size then
        return nil, string.format("Size mismatch after decompression for '%s' (expected %d, got %d)", name, uncomp_size, #content)
      end

      local calc_crc = M.crc32(content)
      if calc_crc ~= crc then
        return nil, string.format("CRC-32 mismatch for '%s' (expected 0x%08x, got 0x%08x)", name, crc, calc_crc)
      end

      files[name] = content
      order[#order + 1] = name
      entries[#entries + 1] = {
        name = name,
        data = content,
        method = method,
        flags = flags,
        dos_time = dos_time,
        dos_date = dos_date,
        crc = crc,
      }
    end
  end

  return {
    files = files,
    order = order,
    entries = entries,
  }
end

function M.read_zip(filepath)
  local f, err = io.open(filepath, "rb")
  if not f then
    return nil, "Cannot open file: " .. tostring(err)
  end
  local bytes = f:read("*a")
  f:close()
  return M.read_zip_bytes(bytes)
end

local function is_precompressed_media(name, size)
  if size < 4096 then
    return false
  end
  local lower = name:lower()
  return lower:match("%.png$") ~= nil
    or lower:match("%.jpe?g$") ~= nil
    or lower:match("%.gif$") ~= nil
    or lower:match("%.webp$") ~= nil
end

-- Build a standard PKZIP byte string from `{ files = { [name] = content }, order = { ... } }`
function M.write_zip_bytes(archive)
  local files = archive.files or {}
  local order = archive.order or {}

  -- Collect all filenames deterministically
  local seen = {}
  local final_order = {}
  for _, name in ipairs(order) do
    if files[name] ~= nil and not seen[name] then
      seen[name] = true
      final_order[#final_order + 1] = name
    end
  end
  local extra_names = {}
  for name, _ in pairs(files) do
    if not seen[name] then
      extra_names[#extra_names + 1] = name
    end
  end
  table.sort(extra_names)
  for _, name in ipairs(extra_names) do
    final_order[#final_order + 1] = name
  end

  local local_parts = {}
  local cd_parts = {}
  local offset = 0
  local count = 0

  -- DOS date/time: 1980-01-01 00:00:00 (date = 0x0021, time = 0x0000)
  local dos_time = 0x0000
  local dos_date = 0x0021

  for _, name in ipairs(final_order) do
    local content = files[name]
    if content ~= nil then
      count = count + 1
      local uncomp_size = #content
      local crc = M.crc32(content)
      local method = 8
      local comp_data
      if is_precompressed_media(name, uncomp_size) then
        method = 0
        comp_data = content
      else
        comp_data = M.deflate(content)
        if #comp_data >= uncomp_size then
          -- Store uncompressed if DEFLATE does not shrink the entry
          method = 0
          comp_data = content
        end
      end
      local comp_size = #comp_data
      local name_len = #name
      local flags = 0x0000 -- Standard flags

      local local_header = table.concat({
        write_u32(SIG_LOCAL_FILE),
        write_u16(20),          -- version needed to extract (2.0)
        write_u16(flags),
        write_u16(method),
        write_u16(dos_time),
        write_u16(dos_date),
        write_u32(crc),
        write_u32(comp_size),
        write_u32(uncomp_size),
        write_u16(name_len),
        write_u16(0),           -- extra field length
        name,
      })

      local cd_header = table.concat({
        write_u32(SIG_CENTRAL_DIR),
        write_u16(20),          -- version made by
        write_u16(20),          -- version needed to extract
        write_u16(flags),
        write_u16(method),
        write_u16(dos_time),
        write_u16(dos_date),
        write_u32(crc),
        write_u32(comp_size),
        write_u32(uncomp_size),
        write_u16(name_len),
        write_u16(0),           -- extra field length
        write_u16(0),           -- file comment length
        write_u16(0),           -- disk number start
        write_u16(0),           -- internal file attributes
        write_u32(0),           -- external file attributes
        write_u32(offset),      -- relative offset of local header
        name,
      })

      local_parts[#local_parts + 1] = local_header
      local_parts[#local_parts + 1] = comp_data
      cd_parts[#cd_parts + 1] = cd_header

      offset = offset + #local_header + comp_size
    end
  end

  local cd_block = table.concat(cd_parts)
  local cd_size = #cd_block
  local cd_offset = offset

  local eocd = table.concat({
    write_u32(SIG_EOCD),
    write_u16(0),               -- number of this disk
    write_u16(0),               -- disk where central directory starts
    write_u16(count),           -- number of central directory records on this disk
    write_u16(count),           -- total number of central directory records
    write_u32(cd_size),
    write_u32(cd_offset),
    write_u16(0),               -- comment length
  })

  return table.concat(local_parts) .. cd_block .. eocd
end

function M.write_zip(filepath, archive)
  local bytes, err = M.write_zip_bytes(archive)
  if not bytes then
    return false, err
  end
  local f, open_err = io.open(filepath, "wb")
  if not f then
    return false, "Cannot open file for writing: " .. tostring(open_err)
  end
  f:write(bytes)
  f:close()
  return true
end

return M
