-- Pure-Lua OOXML (WordprocessingML .docx) reader and writer for nvim-docx.
-- Supports styles preservation, media extraction/embedding, lists, tables,
-- hyperlinks, footnotes, endnotes, code blocks, blockquotes, track changes, and metadata.

local xml = require("nvim-docx.xml")
local image = require("nvim-docx.image")
local markdown = require("nvim-docx.markdown")
local zip = require("nvim-docx.zip")

local M = {}

local sort_rpr_children
local run_attr_elements
local normalize_run_attrs

local NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
local NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
local NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
local NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
local NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
local NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
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
local REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
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

local SYMBOL_FONT_MAP = {
  [0x22] = "\226\136\128", -- ∀
  [0x24] = "\226\136\131", -- ∃
  [0x27] = "\226\136\141", -- ∋
  [0x2A] = "\226\136\151", -- ∗
  [0x2D] = "\226\136\146", -- −
  [0x40] = "\226\137\133", -- ≅
  [0x41] = "\206\145", [0x42] = "\206\146", [0x43] = "\206\167", [0x44] = "\206\148",
  [0x45] = "\206\149", [0x46] = "\206\166", [0x47] = "\206\147", [0x48] = "\206\151",
  [0x49] = "\206\153", [0x4A] = "\207\145", [0x4B] = "\206\154", [0x4C] = "\206\155",
  [0x4D] = "\206\156", [0x4E] = "\206\157", [0x4F] = "\206\159", [0x50] = "\206\160",
  [0x51] = "\206\152", [0x52] = "\206\161", [0x53] = "\206\163", [0x54] = "\206\164",
  [0x55] = "\206\165", [0x56] = "\207\130", [0x57] = "\206\169", [0x58] = "\206\158",
  [0x59] = "\206\168", [0x5A] = "\206\150",
  [0x5C] = "\226\136\180", -- ∴
  [0x5E] = "\226\159\130", -- ⊥
  [0x61] = "\206\177", [0x62] = "\206\178", [0x63] = "\207\135", [0x64] = "\206\180",
  [0x65] = "\206\181", [0x66] = "\207\134", [0x67] = "\206\179", [0x68] = "\206\183",
  [0x69] = "\206\185", [0x6A] = "\207\149", [0x6B] = "\206\186", [0x6C] = "\206\187",
  [0x6D] = "\206\188", [0x6E] = "\206\189", [0x6F] = "\206\191", [0x70] = "\207\128",
  [0x71] = "\206\184", [0x72] = "\207\129", [0x73] = "\207\131", [0x74] = "\207\132",
  [0x75] = "\207\133", [0x76] = "\207\150", [0x77] = "\207\137", [0x78] = "\206\190",
  [0x79] = "\207\136", [0x7A] = "\206\182",
  [0xA3] = "\226\137\164", -- ≤
  [0xA5] = "\226\136\158", -- ∞
  [0xAB] = "\226\134\148", -- ↔
  [0xAC] = "\226\134\144", -- ←
  [0xAD] = "\226\134\145", -- ↑
  [0xAE] = "\226\134\146", -- →
  [0xAF] = "\226\134\147", -- ↓
  [0xB0] = "\194\176",     -- °
  [0xB1] = "\194\177",     -- ±
  [0xB3] = "\226\137\165", -- ≥
  [0xB4] = "\195\151",     -- ×
  [0xB5] = "\226\136\157", -- ∝
  [0xB6] = "\226\136\130", -- ∂
  [0xB7] = "\226\128\162", -- •
  [0xB8] = "\195\183",     -- ÷
  [0xB9] = "\226\137\160", -- ≠
  [0xBA] = "\226\137\161", -- ≡
  [0xBB] = "\226\137\136", -- ≈
  [0xBC] = "\226\128\166", -- …
  [0xC4] = "\226\138\151", -- ⊗
  [0xC5] = "\226\138\149", -- ⊕
  [0xC6] = "\226\136\133", -- ∅
  [0xC7] = "\226\136\169", -- ∩
  [0xC8] = "\226\136\170", -- ∪
  [0xCC] = "\226\138\130", -- ⊂
  [0xCD] = "\226\138\134", -- ⊆
  [0xCE] = "\226\136\136", -- ∈
  [0xCF] = "\226\136\137", -- ∉
  [0xD1] = "\226\136\135", -- ∇
  [0xD2] = "\194\174",     -- ®
  [0xD3] = "\194\169",     -- ©
  [0xD4] = "\226\132\162", -- ™
  [0xD5] = "\226\136\143", -- ∏
  [0xD6] = "\226\136\154", -- √
  [0xD7] = "\226\139\133", -- ⋅
  [0xD8] = "\194\172",     -- ¬
  [0xD9] = "\226\136\167", -- ∧
  [0xDA] = "\226\136\168", -- ∨
  [0xDB] = "\226\135\148", -- ⇔
  [0xDC] = "\226\135\144", -- ⇐
  [0xDD] = "\226\135\145", -- ⇑
  [0xDE] = "\226\135\146", -- ⇒
  [0xDF] = "\226\135\147", -- ⇓
  [0xE5] = "\226\136\145", -- ∑
  [0xF2] = "\226\136\171", -- ∫
}

local function utf8_from_codepoint(cp)
  if not cp or cp < 0 or cp > 0x10FFFF then
    return ""
  end
  if cp <= 0x7F then
    return string.char(cp)
  elseif cp <= 0x7FF then
    return string.char(0xC0 + math.floor(cp / 64), 0x80 + (cp % 64))
  elseif cp <= 0xFFFF then
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

local function decode_w_sym(font, char_hex)
  if not char_hex or char_hex == "" then
    return ""
  end
  local cp = tonumber(char_hex, 16)
  if not cp then
    return ""
  end
  local low = (cp >= 0xF000 and cp <= 0xF0FF) and (cp - 0xF000) or cp
  local fl = (font or ""):lower()
  if fl == "symbol" then
    if SYMBOL_FONT_MAP[low] then
      return SYMBOL_FONT_MAP[low]
    elseif low >= 0x20 and low <= 0x7E then
      return string.char(low)
    end
  end
  return utf8_from_codepoint(cp)
end

local LATEX_NARY_SYMBOLS = {
  ["int"] = "\226\136\171",    -- ∫
  ["iint"] = "\226\136\172",   -- ∬
  ["iiint"] = "\226\136\173",  -- ∭
  ["oint"] = "\226\136\174",   -- ∮
  ["sum"] = "\226\136\145",    -- ∑
  ["prod"] = "\226\136\143",   -- ∏
  ["coprod"] = "\226\136\144", -- ∐
  ["bigcup"] = "\226\139\131", -- ⋃
  ["bigcap"] = "\226\139\130", -- ⋂
}

local LATEX_MATH_SYMBOLS = {
  ["alpha"] = "\206\177",
  ["beta"] = "\206\178",
  ["gamma"] = "\206\179",
  ["delta"] = "\206\180",
  ["epsilon"] = "\207\181",
  ["varepsilon"] = "\206\181",
  ["zeta"] = "\206\182",
  ["eta"] = "\206\183",
  ["theta"] = "\206\184",
  ["vartheta"] = "\207\145",
  ["iota"] = "\206\185",
  ["kappa"] = "\206\186",
  ["lambda"] = "\206\187",
  ["mu"] = "\206\188",
  ["nu"] = "\206\189",
  ["xi"] = "\206\190",
  ["pi"] = "\207\128",
  ["varpi"] = "\207\150",
  ["rho"] = "\207\129",
  ["varrho"] = "\207\177",
  ["sigma"] = "\207\131",
  ["varsigma"] = "\207\130",
  ["tau"] = "\207\132",
  ["upsilon"] = "\207\133",
  ["phi"] = "\207\149",
  ["varphi"] = "\207\134",
  ["chi"] = "\207\135",
  ["psi"] = "\207\136",
  ["omega"] = "\207\137",
  ["Gamma"] = "\206\147",
  ["Delta"] = "\206\148",
  ["Theta"] = "\206\152",
  ["Lambda"] = "\206\155",
  ["Xi"] = "\206\158",
  ["Pi"] = "\206\160",
  ["Sigma"] = "\206\163",
  ["Upsilon"] = "\206\165",
  ["Phi"] = "\206\166",
  ["Psi"] = "\206\168",
  ["Omega"] = "\206\169",
  ["infty"] = "\226\136\158",
  ["partial"] = "\226\136\130",
  ["nabla"] = "\226\136\135",
  ["forall"] = "\226\136\128",
  ["exists"] = "\226\136\131",
  ["emptyset"] = "\226\136\133",
  ["varnothing"] = "\226\136\133",
  ["pm"] = "\194\177",
  ["mp"] = "\226\136\147",
  ["times"] = "\195\151",
  ["div"] = "\195\183",
  ["cdot"] = "\226\139\133",
  ["ast"] = "\226\136\151",
  ["star"] = "\226\139\134",
  ["circ"] = "\226\136\152",
  ["bullet"] = "\226\128\162",
  ["leq"] = "\226\137\164",
  ["le"] = "\226\137\164",
  ["geq"] = "\226\137\165",
  ["ge"] = "\226\137\165",
  ["neq"] = "\226\137\160",
  ["ne"] = "\226\137\160",
  ["approx"] = "\226\137\136",
  ["equiv"] = "\226\137\161",
  ["sim"] = "\226\136\188",
  ["simeq"] = "\226\137\131",
  ["cong"] = "\226\137\133",
  ["propto"] = "\226\136\157",
  ["in"] = "\226\136\136",
  ["notin"] = "\226\136\137",
  ["ni"] = "\226\136\141",
  ["subset"] = "\226\138\130",
  ["supset"] = "\226\138\131",
  ["subseteq"] = "\226\138\134",
  ["supseteq"] = "\226\138\135",
  ["cup"] = "\226\136\170",
  ["cap"] = "\226\136\169",
  ["setminus"] = "\226\136\150",
  ["to"] = "\226\134\146",
  ["rightarrow"] = "\226\134\146",
  ["leftarrow"] = "\226\134\144",
  ["leftrightarrow"] = "\226\134\148",
  ["Rightarrow"] = "\226\135\146",
  ["Leftarrow"] = "\226\135\144",
  ["Leftrightarrow"] = "\226\135\148",
  ["mapsto"] = "\226\134\166",
  ["ldots"] = "\226\128\166",
  ["cdots"] = "\226\139\175",
  ["dots"] = "\226\128\166",
  ["vdots"] = "\226\139\174",
  ["ddots"] = "\226\139\177",
  ["aleph"] = "\226\132\181",
  ["hbar"] = "\226\132\143",
  ["ell"] = "\226\132\147",
  ["Re"] = "\226\132\156",
  ["Im"] = "\226\132\145",
  ["angle"] = "\226\136\160",
  ["perp"] = "\226\159\130",
  ["parallel"] = "\226\136\165",
  ["triangle"] = "\226\150\179",
  ["oplus"] = "\226\138\149",
  ["otimes"] = "\226\138\151",
  ["neg"] = "\194\172",
  ["land"] = "\226\136\167",
  ["lor"] = "\226\136\168",
}

local UNICODE_MATH_TO_LATEX = {
  ["\226\136\171"] = "\\int",
  ["\226\136\172"] = "\\iint",
  ["\226\136\173"] = "\\iiint",
  ["\226\136\174"] = "\\oint",
  ["\226\136\145"] = "\\sum",
  ["\226\136\143"] = "\\prod",
  ["\226\136\144"] = "\\coprod",
  ["\226\139\131"] = "\\bigcup",
  ["\226\139\130"] = "\\bigcap",
  ["\206\177"] = "\\alpha",
  ["\206\178"] = "\\beta",
  ["\206\179"] = "\\gamma",
  ["\206\180"] = "\\delta",
  ["\207\181"] = "\\epsilon",
  ["\206\181"] = "\\varepsilon",
  ["\206\182"] = "\\zeta",
  ["\206\183"] = "\\eta",
  ["\206\184"] = "\\theta",
  ["\207\145"] = "\\vartheta",
  ["\206\185"] = "\\iota",
  ["\206\186"] = "\\kappa",
  ["\206\187"] = "\\lambda",
  ["\206\188"] = "\\mu",
  ["\206\189"] = "\\nu",
  ["\206\190"] = "\\xi",
  ["\207\128"] = "\\pi",
  ["\207\150"] = "\\varpi",
  ["\207\129"] = "\\rho",
  ["\207\177"] = "\\varrho",
  ["\207\131"] = "\\sigma",
  ["\207\130"] = "\\varsigma",
  ["\207\132"] = "\\tau",
  ["\207\133"] = "\\upsilon",
  ["\207\149"] = "\\phi",
  ["\207\134"] = "\\varphi",
  ["\207\135"] = "\\chi",
  ["\207\136"] = "\\psi",
  ["\207\137"] = "\\omega",
  ["\206\147"] = "\\Gamma",
  ["\206\148"] = "\\Delta",
  ["\206\152"] = "\\Theta",
  ["\206\155"] = "\\Lambda",
  ["\206\158"] = "\\Xi",
  ["\206\160"] = "\\Pi",
  ["\206\163"] = "\\Sigma",
  ["\206\165"] = "\\Upsilon",
  ["\206\166"] = "\\Phi",
  ["\206\168"] = "\\Psi",
  ["\206\169"] = "\\Omega",
  ["\226\136\158"] = "\\infty",
  ["\226\136\130"] = "\\partial",
  ["\226\136\135"] = "\\nabla",
  ["\226\136\128"] = "\\forall",
  ["\226\136\131"] = "\\exists",
  ["\226\136\133"] = "\\emptyset",
  ["\194\177"] = "\\pm",
  ["\226\136\147"] = "\\mp",
  ["\195\151"] = "\\times",
  ["\195\183"] = "\\div",
  ["\226\139\133"] = "\\cdot",
  ["\226\136\151"] = "\\ast",
  ["\226\139\134"] = "\\star",
  ["\226\136\152"] = "\\circ",
  ["\226\137\164"] = "\\leq",
  ["\226\137\165"] = "\\geq",
  ["\226\137\160"] = "\\neq",
  ["\226\137\136"] = "\\approx",
  ["\226\137\161"] = "\\equiv",
  ["\226\136\188"] = "\\sim",
  ["\226\137\131"] = "\\simeq",
  ["\226\137\133"] = "\\cong",
  ["\226\136\157"] = "\\propto",
  ["\226\136\136"] = "\\in",
  ["\226\136\137"] = "\\notin",
  ["\226\136\141"] = "\\ni",
  ["\226\138\130"] = "\\subset",
  ["\226\138\131"] = "\\supset",
  ["\226\138\134"] = "\\subseteq",
  ["\226\138\135"] = "\\supseteq",
  ["\226\136\170"] = "\\cup",
  ["\226\136\169"] = "\\cap",
  ["\226\136\150"] = "\\setminus",
  ["\226\134\146"] = "\\to",
  ["\226\134\144"] = "\\leftarrow",
  ["\226\134\148"] = "\\leftrightarrow",
  ["\226\135\146"] = "\\Rightarrow",
  ["\226\135\144"] = "\\Leftarrow",
  ["\226\135\148"] = "\\Leftrightarrow",
  ["\226\134\166"] = "\\mapsto",
  ["\226\128\166"] = "\\ldots",
  ["\226\139\175"] = "\\cdots",
  ["\226\139\174"] = "\\vdots",
  ["\226\139\177"] = "\\ddots",
  ["\226\132\181"] = "\\aleph",
  ["\226\132\143"] = "\\hbar",
  ["\226\132\147"] = "\\ell",
  ["\226\132\156"] = "\\Re",
  ["\226\132\145"] = "\\Im",
  ["\226\136\160"] = "\\angle",
  ["\226\159\130"] = "\\perp",
  ["\226\136\165"] = "\\parallel",
  ["\226\150\179"] = "\\triangle",
  ["\226\138\149"] = "\\oplus",
  ["\226\138\151"] = "\\otimes",
  ["\194\172"] = "\\neg",
  ["\226\136\167"] = "\\land",
  ["\226\136\168"] = "\\lor",
}

local function omml_text_to_latex(s)
  if not s or s == "" then
    return ""
  end
  local out = {}
  local i = 1
  local n = #s
  while i <= n do
    local b = string.byte(s, i)
    local clen = 1
    if b >= 0xF0 then
      clen = 4
    elseif b >= 0xE0 then
      clen = 3
    elseif b >= 0xC0 then
      clen = 2
    end
    local ch = s:sub(i, i + clen - 1)
    local next_i = i + clen
    local tex = UNICODE_MATH_TO_LATEX[ch]
    if tex then
      local next_ch = (next_i <= n) and s:sub(next_i, next_i) or ""
      if next_ch:match("^[A-Za-z]$") then
        out[#out + 1] = tex .. " "
      else
        out[#out + 1] = tex
      end
    else
      out[#out + 1] = ch
    end
    i = next_i
  end
  return table.concat(out)
end

local function read_math_braced(s, pos)
  if s:sub(pos, pos) ~= "{" then
    return nil, pos
  end
  local depth = 1
  local i = pos + 1
  local n = #s
  while i <= n do
    local c = s:sub(i, i)
    if c == "\\" then
      i = i + 2
    elseif c == "{" then
      depth = depth + 1
      i = i + 1
    elseif c == "}" then
      depth = depth - 1
      if depth == 0 then
        return s:sub(pos + 1, i - 1), i + 1
      end
      i = i + 1
    else
      i = i + 1
    end
  end
  return s:sub(pos + 1), n + 1
end

local function read_math_bracket(s, pos)
  if s:sub(pos, pos) ~= "[" then
    return nil, pos
  end
  local depth = 1
  local i = pos + 1
  local n = #s
  while i <= n do
    local c = s:sub(i, i)
    if c == "[" then
      depth = depth + 1
    elseif c == "]" then
      depth = depth - 1
      if depth == 0 then
        return s:sub(pos + 1, i - 1), i + 1
      end
    end
    i = i + 1
  end
  return s:sub(pos + 1), n + 1
end

local function read_math_script_arg(s, pos)
  local n = #s
  while pos <= n and s:sub(pos, pos):match("%s") do
    pos = pos + 1
  end
  if pos > n then
    return "", pos
  end
  local c = s:sub(pos, pos)
  if c == "{" then
    return read_math_braced(s, pos)
  elseif c == "\\" then
    local _, e_idx, cmd = s:find("^\\([A-Za-z]+)", pos)
    if cmd then
      local next_p = e_idx + 1
      if cmd == "frac" then
        while next_p <= n and s:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
        local g1, p1 = read_math_braced(s, next_p)
        if g1 then
          while p1 <= n and s:sub(p1, p1):match("%s") do p1 = p1 + 1 end
          local g2, p2 = read_math_braced(s, p1)
          if g2 then
            return "\\frac{" .. g1 .. "}{" .. g2 .. "}", p2
          end
        end
      elseif cmd == "sqrt" then
        while next_p <= n and s:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
        local deg_str = nil
        if s:sub(next_p, next_p) == "[" then
          deg_str, next_p = read_math_bracket(s, next_p)
          while next_p <= n and s:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
        end
        local g1, p1 = read_math_braced(s, next_p)
        if g1 then
          if deg_str then
            return "\\sqrt[" .. deg_str .. "]{" .. g1 .. "}", p1
          end
          return "\\sqrt{" .. g1 .. "}", p1
        end
      end
      return "\\" .. cmd, next_p
    end
    return s:sub(pos, math.min(pos + 1, n)), math.min(pos + 2, n + 1)
  else
    return c, pos + 1
  end
end

local function make_omath_run(text)
  return xml.elem("m:r", nil, {
    xml.elem("m:t", { ["xml:space"] = "preserve" }, { xml.text_node(text) }),
  })
end

local function latex_to_omml_nodes(tex)
  if not tex or tex == "" then
    return {}
  end

  local atoms = {}
  local i = 1
  local n = #tex

  while i <= n do
    local c = tex:sub(i, i)
    if c == "\\" then
      local _, e_idx, cmd = tex:find("^\\([A-Za-z]+)", i)
      if cmd then
        local next_p = e_idx + 1
        if cmd == "frac" then
          while next_p <= n and tex:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
          local num_str, p1 = read_math_braced(tex, next_p)
          if not num_str then
            num_str, p1 = read_math_script_arg(tex, next_p)
          end
          while p1 <= n and tex:sub(p1, p1):match("%s") do p1 = p1 + 1 end
          local den_str, p2 = read_math_braced(tex, p1)
          if not den_str then
            den_str, p2 = read_math_script_arg(tex, p1)
          end
          local num_nodes = latex_to_omml_nodes(num_str or "")
          local den_nodes = latex_to_omml_nodes(den_str or "")
          if #num_nodes == 0 then num_nodes = { make_omath_run("") } end
          if #den_nodes == 0 then den_nodes = { make_omath_run("") } end
          atoms[#atoms + 1] = {
            kind = "elem",
            node = xml.elem("m:f", nil, {
              xml.elem("m:num", nil, num_nodes),
              xml.elem("m:den", nil, den_nodes),
            }),
          }
          i = p2
        elseif cmd == "sqrt" then
          while next_p <= n and tex:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
          local deg_str = nil
          if tex:sub(next_p, next_p) == "[" then
            deg_str, next_p = read_math_bracket(tex, next_p)
            while next_p <= n and tex:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
          end
          local body_str, p1 = read_math_braced(tex, next_p)
          if not body_str then
            body_str, p1 = read_math_script_arg(tex, next_p)
          end
          local body_nodes = latex_to_omml_nodes(body_str or "")
          if #body_nodes == 0 then body_nodes = { make_omath_run("") } end
          if deg_str and deg_str ~= "" then
            atoms[#atoms + 1] = {
              kind = "elem",
              node = xml.elem("m:rad", nil, {
                xml.elem("m:deg", nil, latex_to_omml_nodes(deg_str)),
                xml.elem("m:e", nil, body_nodes),
              }),
            }
          else
            atoms[#atoms + 1] = {
              kind = "elem",
              node = xml.elem("m:rad", nil, {
                xml.elem("m:radPr", nil, { xml.elem("m:degHide", { ["m:val"] = "1" }) }),
                xml.elem("m:deg"),
                xml.elem("m:e", nil, body_nodes),
              }),
            }
          end
          i = p1
        elseif LATEX_NARY_SYMBOLS[cmd] then
          local nary_chr = LATEX_NARY_SYMBOLS[cmd]
          local p = next_p
          local sub_str, sup_str = nil, nil
          for _ = 1, 2 do
            local sp = p
            while sp <= n and tex:sub(sp, sp):match("%s") do sp = sp + 1 end
            if sp <= n and tex:sub(sp, sp) == "_" and not sub_str then
              sub_str, p = read_math_script_arg(tex, sp + 1)
            elseif sp <= n and tex:sub(sp, sp) == "^" and not sup_str then
              sup_str, p = read_math_script_arg(tex, sp + 1)
            else
              break
            end
          end
          while p <= n and tex:sub(p, p):match("%s") do p = p + 1 end
          -- Scan operand up to top-level '=', '<', '>' or end of expression
          local op_start = p
          local depth = 0
          local op_end = n
          local scan = p
          while scan <= n do
            local sc = tex:sub(scan, scan)
            if sc == "\\" then
              local _, ce, ccmd = tex:find("^\\([A-Za-z]+)", scan)
              if ccmd then
                if depth == 0 and (ccmd == "leq" or ccmd == "le" or ccmd == "geq" or ccmd == "ge" or ccmd == "neq" or ccmd == "ne" or ccmd == "approx" or ccmd == "equiv" or ccmd == "to" or ccmd == "rightarrow") then
                  op_end = scan - 1
                  break
                end
                scan = ce + 1
              else
                scan = scan + 2
              end
            elseif sc == "{" then
              depth = depth + 1
              scan = scan + 1
            elseif sc == "}" then
              if depth > 0 then depth = depth - 1 end
              scan = scan + 1
            elseif depth == 0 and (sc == "=" or sc == "<" or sc == ">") then
              op_end = scan - 1
              break
            else
              scan = scan + 1
            end
          end
          while op_end >= op_start and tex:sub(op_end, op_end):match("%s") do
            op_end = op_end - 1
          end
          local operand_str = tex:sub(op_start, op_end)
          local nary_pr_children = {
            xml.elem("m:chr", { ["m:val"] = nary_chr }),
            xml.elem("m:limLoc", { ["m:val"] = "subSup" }),
          }
          if not sub_str then
            nary_pr_children[#nary_pr_children + 1] = xml.elem("m:subHide", { ["m:val"] = "1" })
          end
          if not sup_str then
            nary_pr_children[#nary_pr_children + 1] = xml.elem("m:supHide", { ["m:val"] = "1" })
          end
          local e_nodes = latex_to_omml_nodes(operand_str)
          if #e_nodes == 0 then e_nodes = { make_omath_run("") } end
          atoms[#atoms + 1] = {
            kind = "elem",
            node = xml.elem("m:nary", nil, {
              xml.elem("m:naryPr", nil, nary_pr_children),
              xml.elem("m:sub", nil, sub_str and latex_to_omml_nodes(sub_str) or {}),
              xml.elem("m:sup", nil, sup_str and latex_to_omml_nodes(sup_str) or {}),
              xml.elem("m:e", nil, e_nodes),
            }),
          }
          i = op_end + 1
        elseif LATEX_MATH_SYMBOLS[cmd] then
          local sym = LATEX_MATH_SYMBOLS[cmd]
          -- Skip single space after command if followed by letter
          if next_p <= n and tex:sub(next_p, next_p) == " " and tex:sub(next_p + 1, next_p + 1):match("^[A-Za-z]$") then
            next_p = next_p + 1
          end
          atoms[#atoms + 1] = { kind = "text", text = sym }
          i = next_p
        elseif cmd == "left" or cmd == "right" then
          i = next_p
        elseif cmd == "text" or cmd == "mathrm" or cmd == "mathbf" or cmd == "mathit" then
          while next_p <= n and tex:sub(next_p, next_p):match("%s") do next_p = next_p + 1 end
          local grp, p1 = read_math_braced(tex, next_p)
          if grp then
            atoms[#atoms + 1] = { kind = "text", text = grp }
            i = p1
          else
            i = next_p
          end
        else
          atoms[#atoms + 1] = { kind = "text", text = "\\" .. cmd }
          i = next_p
        end
      else
        -- Escaped punctuation like \{, \}, \, , \;
        local esc = tex:sub(i + 1, i + 1)
        if esc == "," or esc == ";" or esc == ":" or esc == "!" or esc == " " then
          atoms[#atoms + 1] = { kind = "text", text = " " }
        elseif esc ~= "" then
          atoms[#atoms + 1] = { kind = "text", text = esc }
        else
          atoms[#atoms + 1] = { kind = "text", text = "\\" }
        end
        i = i + 2
      end
    elseif c == "_" or c == "^" then
      local sub_str, sup_str = nil, nil
      local p = i
      for _ = 1, 2 do
        local sp = p
        while sp <= n and tex:sub(sp, sp):match("%s") do sp = sp + 1 end
        if sp <= n and tex:sub(sp, sp) == "_" and not sub_str then
          sub_str, p = read_math_script_arg(tex, sp + 1)
        elseif sp <= n and tex:sub(sp, sp) == "^" and not sup_str then
          sup_str, p = read_math_script_arg(tex, sp + 1)
        else
          break
        end
      end
      -- Pop preceding non-whitespace atom as base
      local base_nodes = nil
      if #atoms > 0 and not (atoms[#atoms].kind == "text" and atoms[#atoms].text:match("^%s+$")) then
        local last = table.remove(atoms)
        if last.kind == "text" then
          base_nodes = { make_omath_run(last.text) }
        elseif last.kind == "group" then
          base_nodes = last.nodes
        else
          base_nodes = { last.node }
        end
      else
        base_nodes = { make_omath_run("") }
      end

      if sub_str and sup_str then
        atoms[#atoms + 1] = {
          kind = "elem",
          node = xml.elem("m:sSubSup", nil, {
            xml.elem("m:e", nil, base_nodes),
            xml.elem("m:sub", nil, latex_to_omml_nodes(sub_str)),
            xml.elem("m:sup", nil, latex_to_omml_nodes(sup_str)),
          }),
        }
      elseif sup_str then
        atoms[#atoms + 1] = {
          kind = "elem",
          node = xml.elem("m:sSup", nil, {
            xml.elem("m:e", nil, base_nodes),
            xml.elem("m:sup", nil, latex_to_omml_nodes(sup_str)),
          }),
        }
      elseif sub_str then
        atoms[#atoms + 1] = {
          kind = "elem",
          node = xml.elem("m:sSub", nil, {
            xml.elem("m:e", nil, base_nodes),
            xml.elem("m:sub", nil, latex_to_omml_nodes(sub_str)),
          }),
        }
      end
      i = p
    elseif c == "{" then
      local grp, next_p = read_math_braced(tex, i)
      atoms[#atoms + 1] = { kind = "group", nodes = latex_to_omml_nodes(grp or "") }
      i = next_p
    else
      atoms[#atoms + 1] = { kind = "text", text = c }
      i = i + 1
    end
  end

  -- Merge adjacent text atoms into <m:r><m:t> runs
  local result = {}
  local text_buf = {}
  local function flush_text()
    if #text_buf > 0 then
      result[#result + 1] = make_omath_run(table.concat(text_buf))
      text_buf = {}
    end
  end

  for _, a in ipairs(atoms) do
    if a.kind == "text" then
      text_buf[#text_buf + 1] = a.text
    elseif a.kind == "group" then
      flush_text()
      for _, gn in ipairs(a.nodes or {}) do
        result[#result + 1] = gn
      end
    else
      flush_text()
      result[#result + 1] = a.node
    end
  end
  flush_text()

  if #result == 0 then
    result[1] = make_omath_run(tex)
  end
  return result
end

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

local RUN_ATTRS = { "font", "size", "color", "highlight" }

local function each_text_node(inlines, fn)
  for _, it in ipairs(inlines or {}) do
    if it.type == "text" then
      fn(it)
    elseif it.inlines then
      each_text_node(it.inlines, fn)
    end
  end
end

-- Keep font/size/color/highlight only on runs that differ from the paragraph's dominant look (by
-- amount of text); the dominant look is the paragraph's own and is inherited when it is edited
normalize_run_attrs = function(inlines)
  local weight, by_sig = {}, {}
  each_text_node(inlines, function(it)
    local sig = table.concat({ tostring(it.font), tostring(it.size), tostring(it.color), tostring(it.highlight) }, "|")
    weight[sig] = (weight[sig] or 0) + #(it.text or "")
    by_sig[sig] = it
  end)
  local best, best_w = nil, -1
  for sig, w in pairs(weight) do
    if w > best_w then
      best, best_w = sig, w
    end
  end
  local dom = best and by_sig[best] or {}
  local dominant = {}
  for _, k in ipairs(RUN_ATTRS) do
    dominant[k] = dom[k]
  end
  each_text_node(inlines, function(it)
    for _, k in ipairs(RUN_ATTRS) do
      if it[k] == dominant[k] then
        it[k] = nil
      elseif it[k] == nil and dominant[k] ~= nil then
        -- a run without the attribute inside a paragraph that mostly has it: leave as is
        it[k] = nil
      end
    end
  end)
  return inlines
end

-- Map a character in a symbol font (Symbol, Wingdings, ...) to Unicode; char_hex is its code in hex
function M.decode_symbol(font, char_hex)
  return decode_w_sym(font, char_hex)
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
          if workspace and workspace ~= "" and rawget(archive.files, zip_entry_path) then
            write_binary_file(workspace .. "/" .. rel_out_path, archive.files[zip_entry_path])
          end
        end
      end
    end
  end
  extract_rels_media(doc_rels)
  extract_rels_media(footnote_rels)
  extract_rels_media(endnote_rels)

  -- Also extract any other files in word/media/* into workspace/media/* (large ones are left
  -- compressed in the archive until needed; see M.extract_media)
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
      return omml_text_to_latex(xml.get_direct_text(mnode))
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
    elseif lname == "nary" then
      local npr = xml.find_child(mnode, "naryPr")
      local chr_node = npr and xml.find_child(npr, "chr") or nil
      local chr_val = (chr_node and (xml.get_attr(chr_node, "m:val") or xml.get_attr(chr_node, "val"))) or "\226\136\171"
      local chr_tex = UNICODE_MATH_TO_LATEX[chr_val] or omml_text_to_latex(chr_val)
      local sub = extract_omath_text(xml.find_child(mnode, "sub"))
      local sup = extract_omath_text(xml.find_child(mnode, "sup"))
      local body = extract_omath_text(xml.find_child(mnode, "e"))
      local res = chr_tex
      if sub ~= "" then
        res = res .. "_{" .. sub .. "}"
      end
      if sup ~= "" then
        res = res .. "^{" .. sup .. "}"
      end
      if body ~= "" then
        if not body:match("^%s") then
          res = res .. " " .. body
        else
          res = res .. body
        end
      end
      return res
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
    local is_chart = false
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
      for _, ch in ipairs(xml.find_all(drawing_node, "chart")) do
        r_embed = xml.get_attr(ch, "r:id")
        if r_embed then
          is_chart = true
          break
        end
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
    if is_chart and alt == "" then
      for _, doc_pr in ipairs(xml.find_all(drawing_node, "docPr")) do
        local nm = xml.get_attr(doc_pr, "name")
        if nm and nm ~= "" then
          alt = nm
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
    local page_break = nil
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

              -- explicit font / size / color / highlight (shown as attribute spans when they differ
              -- from the paragraph's dominant look, see normalize_run_attrs)
              local rf = xml.find_child(rpr, "rFonts")
              local rf_name = rf and (xml.get_attr(rf, "w:ascii") or xml.get_attr(rf, "w:hAnsi")) or nil
              if rf_name and not MONOSPACE_FONTS[rf_name:lower()] then
                rfmt.font = rf_name
              end
              local sz_node = xml.find_child(rpr, "sz")
              local sz = sz_node and tonumber(xml.get_attr(sz_node, "w:val"))
              if sz then
                rfmt.size = sz / 2
              end
              local color_node = xml.find_child(rpr, "color")
              local cval = color_node and xml.get_attr(color_node, "w:val")
              if cval and cval:match("^%x%x%x%x%x%x$") and cval:upper() ~= "000000" then
                rfmt.color = "#" .. cval:upper()
              end
              local hl_node = xml.find_child(rpr, "highlight")
              local hval = hl_node and xml.get_attr(hl_node, "w:val")
              if hval and hval ~= "none" then
                rfmt.highlight = hval
              end

              local u_node = xml.find_child(rpr, "u")
              if u_node then
                local uval = (xml.get_attr(u_node, "w:val") or "single"):lower()
                rfmt.underline = uval ~= "none" and uval ~= "0" and uval ~= "false"
              end

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

          local function walk_r_children(rchildren)
            for _, rchild in ipairs(rchildren or {}) do
              local rlname = rchild.local_name
              if rlname == "AlternateContent" then
                local branch = xml.find_child(rchild, "Choice") or xml.find_child(rchild, "Fallback")
                if branch then
                  walk_r_children(branch.children)
                end
              elseif rlname == "fldChar" then
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
              elseif rlname == "sym" then
                local sym_font = xml.get_attr(rchild, "w:font")
                local sym_char = xml.get_attr(rchild, "w:char")
                local sym_txt = decode_w_sym(sym_font, sym_char)
                if sym_txt ~= "" then
                  local node = vim.deepcopy(rfmt)
                  node.type = "text"
                  node.text = sym_txt
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
                  page_break = (#inlines > 0) and "after" or "before"
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
              elseif rlname == "softHyphen" then
                local node = vim.deepcopy(rfmt)
                node.type = "text"
                node.text = "\194\173"
                push_inline(node)
              elseif rlname == "drawing" or rlname == "pict" or rlname == "object" then
                local txbx_list = xml.find_all(rchild, "txbxContent")
                local dgm = xml.find_all(rchild, "relIds")[1]
                if #txbx_list > 0 then
                  for _, txbx in ipairs(txbx_list) do
                    for _, tp in ipairs(xml.find_children(txbx, "p")) do
                      local tpinl = extract_inlines_from_nodes(tp.children, false, nil, active_rels)
                      if #tpinl > 0 then
                        if #inlines > 0 and inlines[#inlines].type ~= "linebreak" then
                          push_inline({ type = "linebreak" })
                        end
                        for _, item in ipairs(tpinl) do
                          push_inline(item)
                        end
                      end
                    end
                  end
                elseif dgm then
                  local r_dm = xml.get_attr(dgm, "r:dm")
                  local rel = r_dm and ((active_rels and active_rels[r_dm]) or doc_rels[r_dm]) or nil
                  if rel and rel.target and rel.target ~= "" then
                    local dgm_path = normalize_zip_path(doc_dir, rel.target)
                    local dgm_raw = archive.files[dgm_path]
                    local dgm_root = dgm_raw and xml.parse(dgm_raw) or nil
                    if dgm_root then
                      for _, pt in ipairs(xml.find_all(dgm_root, "pt")) do
                        local dt = xml.find_child(pt, "t")
                        if dt then
                          for _, ap in ipairs(xml.find_children(dt, "p")) do
                            local parts = {}
                            for _, at in ipairs(xml.find_all(ap, "t")) do
                              parts[#parts + 1] = xml.get_direct_text(at)
                            end
                            local ltxt = table.concat(parts):gsub("^%s+", ""):gsub("%s+$", "")
                            if ltxt ~= "" then
                              if #inlines > 0 and inlines[#inlines].type ~= "linebreak" then
                                push_inline({ type = "linebreak" })
                              end
                              push_inline({ type = "text", text = ltxt })
                            end
                          end
                        end
                      end
                    end
                  end
                else
                  local d = extract_run_drawing(rchild, active_rels)
                  if d then
                    if d.is_hr then
                      found_hr = true
                    else
                      push_inline(d)
                    end
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
          end

          walk_r_children(child.children)

        elseif lname == "AlternateContent" then
          local branch = xml.find_child(child, "Choice") or xml.find_child(child, "Fallback")
          if branch then
            walk_inline_children(branch.children, cur_fmt, in_del)
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
    return markdown.merge_inlines(inlines), found_hr, found_checkbox, only_omath_para, page_break
  end

  local trim_inlines = markdown.trim_inlines

  local function collect_paragraphs_recursive(container, out_paras)
    out_paras = out_paras or {}
    for _, c in ipairs((container and container.children) or {}) do
      local cl = c.local_name
      if cl == "p" then
        out_paras[#out_paras + 1] = c
      elseif cl == "sdt" then
        local sc = xml.find_child(c, "sdtContent")
        if sc then collect_paragraphs_recursive(sc, out_paras) end
      elseif cl == "customXml" or cl == "smartTag" then
        collect_paragraphs_recursive(c, out_paras)
      elseif cl == "AlternateContent" then
        local br = xml.find_child(c, "Choice") or xml.find_child(c, "Fallback")
        if br then collect_paragraphs_recursive(br, out_paras) end
      elseif cl == "ins" or cl == "moveTo" then
        if track_policy ~= "reject" then
          collect_paragraphs_recursive(c, out_paras)
        end
      elseif cl == "del" or cl == "moveFrom" then
        if track_policy == "reject" or track_policy == "all" then
          collect_paragraphs_recursive(c, out_paras)
        end
      end
    end
    return out_paras
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
        local paras = collect_paragraphs_recursive(fnode)
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
  local merged_into_prev_block = false
  local page_break_after = false

  local function process_paragraph_inner(pnode)
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
      local btm = pbdr and xml.find_child(pbdr, "bottom") or nil
      local btm_val = btm and (xml.get_attr(btm, "w:val") or ""):lower() or nil
      if btm_val and btm_val ~= "" and btm_val ~= "none" and btm_val ~= "nil" then
        has_bottom_border = true
      end
    end

    local is_code = is_code_para_style(styles, pstyle)
    local inlines, found_hr, found_checkbox, only_omath_para, page_break = extract_inlines_from_nodes(pnode.children, is_code)
    local ppr_node = xml.find_child(pnode, "pPr")
    if not page_break and ppr_node and is_bool_on(xml.find_child(ppr_node, "pageBreakBefore")) then
      page_break = "before"
    end
    if page_break == "before" then
      ast.blocks[#ast.blocks + 1] = { type = "pagebreak" }
    elseif page_break == "after" then
      page_break_after = true
    end

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
        merged_into_prev_block = true
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
    normalize_run_attrs(inlines)

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

  -- Remember the source XML node of every block so a save can reuse it (see preserve.lua)
  local function process_paragraph(pnode)
    local n0 = #ast.blocks
    page_break_after = false
    process_paragraph_inner(pnode)
    if page_break_after then
      ast.blocks[#ast.blocks + 1] = { type = "pagebreak" }
    end
    for k = n0 + 1, #ast.blocks do
      ast.blocks[k]._pnode = pnode
    end
  end

  local function process_table(tbl_node)
    ordered_counters = {}
    seen_override_nums = {}
    local all_rows = {}
    local alignments = {}

    local function collect_tc(tr_container, out_tcs)
      out_tcs = out_tcs or {}
      for _, c in ipairs((tr_container and tr_container.children) or {}) do
        local cl = c.local_name
        if cl == "tc" then
          out_tcs[#out_tcs + 1] = c
        elseif cl == "sdt" then
          local sc = xml.find_child(c, "sdtContent")
          if sc then collect_tc(sc, out_tcs) end
        elseif cl == "customXml" or cl == "smartTag" then
          collect_tc(c, out_tcs)
        elseif cl == "AlternateContent" then
          local br = xml.find_child(c, "Choice") or xml.find_child(c, "Fallback")
          if br then collect_tc(br, out_tcs) end
        elseif cl == "ins" or cl == "moveTo" then
          if track_policy ~= "reject" then
            collect_tc(c, out_tcs)
          end
        elseif cl == "del" or cl == "moveFrom" then
          if track_policy == "reject" or track_policy == "all" then
            collect_tc(c, out_tcs)
          end
        end
      end
      return out_tcs
    end

    local function collect_tr(container)
      for _, child in ipairs(container.children or {}) do
        local cl = child.local_name
        if cl == "tr" then
          local row_cells = {}
          local col_idx = 0
          for _, tc in ipairs(collect_tc(child)) do
            col_idx = col_idx + 1
            local cell_inlines = {}
            local paras = collect_paragraphs_recursive(tc)
            for _, p in ipairs(paras) do
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
              normalize_run_attrs(pinl)
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
          if #row_cells > 0 then
            all_rows[#all_rows + 1] = row_cells
          end
        elseif cl == "sdt" then
          local sdt_c = xml.find_child(child, "sdtContent")
          if sdt_c then collect_tr(sdt_c) end
        elseif cl == "customXml" or cl == "smartTag" then
          collect_tr(child)
        elseif cl == "AlternateContent" then
          local br = xml.find_child(child, "Choice") or xml.find_child(child, "Fallback")
          if br then collect_tr(br) end
        elseif cl == "ins" or cl == "moveTo" then
          if track_policy ~= "reject" then
            collect_tr(child)
          end
        elseif cl == "del" or cl == "moveFrom" then
          if track_policy == "reject" or track_policy == "all" then
            collect_tr(child)
          end
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
      _pnode = tbl_node,
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
      elseif lname == "AlternateContent" then
        local branch = xml.find_child(child, "Choice") or xml.find_child(child, "Fallback")
        if branch then
          walk_body(branch)
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

  -- Walk top-level body children one by one, grouping the blocks each one produces.
  -- Children that produce no block (empty spacing paragraphs, page/section breaks, bookmarks)
  -- are kept as "glue" in front of the next group so a save can put them back.
  local groups = {}
  local glue = {}
  for _, child in ipairs(body.children or {}) do
    if child.tag and child.local_name ~= "sectPr" then
      local n0 = #ast.blocks
      merged_into_prev_block = false
      walk_body({ children = { child } })
      local n1 = #ast.blocks
      if n1 > n0 then
        local group = { first = n0 + 1, last = n1, nodes = { child }, glue = glue }
        groups[#groups + 1] = group
        for k = n0 + 1, n1 do
          ast.blocks[k]._group = #groups
        end
        glue = {}
      elseif merged_into_prev_block and #groups > 0 and groups[#groups].last == n1 then
        local nodes = groups[#groups].nodes
        nodes[#nodes + 1] = child
      else
        glue[#glue + 1] = child
      end
    end
  end
  ast._groups = groups
  ast._trailing_glue = glue

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
    '<w:style w:type="paragraph" w:styleId="Title">',
      '<w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:spacing w:after="80"/><w:contextualSpacing/></w:pPr>',
      '<w:rPr><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Subtitle">',
      '<w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:rPr><w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>',
    '</w:style>',
    '<w:style w:type="paragraph" w:styleId="Quote">',
      '<w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>',
      '<w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>',
      '<w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>',
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
    return ensure_required_styles(default_styles_xml())
  end
  local root = xml.parse(styles_xml_str)
  if not root then
    return ensure_required_styles(default_styles_xml())
  end

  local existing = {}
  local style_nodes = {}
  local heading_ids = {}
  local name_to_id = {}
  for _, snode in ipairs(xml.find_children(root, "style")) do
    local sid = xml.get_attr(snode, "w:styleId")
    if sid then
      existing[sid] = true
      style_nodes[sid] = snode
      local name_node = xml.find_child(snode, "name")
      local sname = (name_node and xml.get_attr(name_node, "w:val") or sid):lower()
      name_to_id[sname] = sid
      name_to_id[sid:lower()] = name_to_id[sid:lower()] or sid
      local h_lvl = sname:match("^heading%s*([1-6])$") or sid:lower():match("^heading([1-6])$")
      if h_lvl then
        heading_ids[tonumber(h_lvl)] = sid
      end
    end
  end

  -- LibreOffice templates leave Normal without spacing and put body-text spacing on TextBody; new
  -- paragraphs use TextBody there so they match the rest of the document. Existing styles are never
  -- modified, since every untouched paragraph depends on them.
  local body_style_id = nil
  local normal_node = style_nodes["Normal"]
  local textbody_node = style_nodes["TextBody"]
  if normal_node and textbody_node then
    local norm_ppr = xml.find_child(normal_node, "pPr")
    local tb_ppr = xml.find_child(textbody_node, "pPr")
    if tb_ppr and xml.find_child(tb_ppr, "spacing") and not (norm_ppr and xml.find_child(norm_ppr, "spacing")) then
      body_style_id = "TextBody"
    end
  end
  local modified = false

  local def_root = xml.parse(default_styles_xml())
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
        local nn = xml.find_child(def_style, "name")
        name_to_id[(nn and xml.get_attr(nn, "w:val") or sid):lower()] = sid
        name_to_id[sid:lower()] = sid
        modified = true
      end
    end
  end

  if modified then
    return xml.serialize(root), heading_ids, body_style_id, name_to_id
  end
  return styles_xml_str, heading_ids, body_style_id, name_to_id
end

local function build_numbering_xml(list_defs, existing_numbering_xml)
  local abs_children = {}
  local num_children = {}
  local root_attrs = {
    ["xmlns:w"] = NS_W,
    ["xmlns:r"] = NS_R,
  }

  if existing_numbering_xml and existing_numbering_xml ~= "" then
    local ex_root = xml.parse(existing_numbering_xml)
    if ex_root then
      for k, v in pairs(ex_root.attrs or {}) do
        root_attrs[k] = v
      end
      root_attrs["xmlns:w"] = NS_W
      root_attrs["xmlns:r"] = NS_R
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

  local root = xml.elem("w:numbering", root_attrs, children)
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

local function build_chart_drawing_element(r_id, chart_id, name, cx, cy)
  cx = tostring(math.floor(tonumber(cx) or 4099560))
  cy = tostring(math.floor(tonumber(cy) or 2059940))
  chart_id = tostring(chart_id or 1)
  name = name or ""

  return xml.elem("w:drawing", nil, {
    xml.elem("wp:inline", {
      ["distT"] = "0",
      ["distB"] = "0",
      ["distL"] = "0",
      ["distR"] = "0",
    }, {
      xml.elem("wp:extent", { ["cx"] = cx, ["cy"] = cy }),
      xml.elem("wp:effectExtent", { ["l"] = "0", ["t"] = "0", ["r"] = "0", ["b"] = "0" }),
      xml.elem("wp:docPr", { ["id"] = chart_id, ["name"] = name }),
      xml.elem("wp:cNvGraphicFramePr"),
      xml.elem("a:graphic", { ["xmlns:a"] = NS_A }, {
        xml.elem("a:graphicData", { ["uri"] = NS_C }, {
          xml.elem("c:chart", {
            ["xmlns:c"] = NS_C,
            ["xmlns:r"] = NS_R,
            ["r:id"] = r_id,
          }),
        }),
      }),
    }),
  })
end

-- Schema order of <w:rPr> children (CT_RPr); Word rejects runs whose properties are out of order
local RPR_ORDER = {}
for idx, name in ipairs({
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
  "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
  "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
  "em", "lang", "eastAsianLayout", "specVanish", "oMath",
}) do
  RPR_ORDER[name] = idx
end

sort_rpr_children = function(children)
  local indexed = {}
  for pos, c in ipairs(children) do
    indexed[pos] = { node = c, pos = pos, rank = RPR_ORDER[c.local_name] or 1000 }
  end
  table.sort(indexed, function(x, y)
    if x.rank ~= y.rank then
      return x.rank < y.rank
    end
    return x.pos < y.pos
  end)
  for pos, e in ipairs(indexed) do
    children[pos] = e.node
  end
end

-- Properties Markdown expresses itself; an edited block takes them from the Markdown, not the original run
local MARKDOWN_RPR = {
  b = true, bCs = true, i = true, iCs = true, strike = true, dstrike = true, u = true, vertAlign = true,
  rPrChange = true,
}
local MARKDOWN_RSTYLES = { Hyperlink = true, VerbatimChar = true, FootnoteReference = true }

-- rPr of the first text run of an original paragraph, minus what Markdown controls
local function inheritable_run_props(pnode)
  if not pnode then
    return nil
  end
  -- the run formatting covering the most text is the paragraph's own look
  local found = nil
  local weights, rprs = {}, {}
  local function walk(node)
    for _, c in ipairs(node.children or {}) do
      if c.local_name == "r" then
        local t = xml.find_child(c, "t")
        if t then
          local rpr = xml.find_child(c, "rPr")
          local key = rpr and xml.serialize(rpr, { declaration = false }) or ""
          weights[key] = (weights[key] or 0) + #(xml.get_all_text(t) or "")
          rprs[key] = rpr or false
        end
      elseif c.local_name == "hyperlink" or c.local_name == "ins" or c.local_name == "smartTag"
        or c.local_name == "customXml" or c.local_name == "sdt" or c.local_name == "sdtContent"
        or c.local_name == "fldSimple" then
        walk(c)
      end
    end
  end
  walk(pnode)
  local best_w = -1
  for key, w in pairs(weights) do
    if w > best_w then
      best_w = w
      found = rprs[key]
    end
  end
  if not found then
    return nil
  end
  local out = {}
  for _, c in ipairs(found.children or {}) do
    if c.tag and not MARKDOWN_RPR[c.local_name] then
      local style = c.local_name == "rStyle" and xml.get_attr(c, "w:val") or nil
      if not (style and MARKDOWN_RSTYLES[style]) then
        out[#out + 1] = c
      end
    end
  end
  return #out > 0 and out or nil
end

local function inheritable_para_props(pnode)
  local ppr = pnode and pnode.local_name == "p" and xml.find_child(pnode, "pPr") or nil
  if not ppr then
    return nil
  end
  local copy = xml.clone(ppr)
  -- Paragraph-mark revision marks belong to the old text
  for k = #copy.children, 1, -1 do
    local c = copy.children[k]
    if c.local_name == "pPrChange" then
      table.remove(copy.children, k)
    end
  end
  return copy
end

-- Relationship ids (r:id, r:embed, r:link, r:dm, ...) referenced anywhere inside a node
local function collect_rel_ids(node, acc)
  local function scan(n)
    for k, v in pairs(n.attrs or {}) do
      if k:sub(1, 2) == "r:" then
        acc[v] = true
      end
    end
  end
  if node.tag then
    scan(node)
  end
  xml.walk(node, scan)
  return acc
end

local NOTE_REFS = { footnoteReference = true, endnoteReference = true, commentReference = true }

local function node_has_note_refs(node)
  return xml.walk(node, function(c)
    return NOTE_REFS[c.local_name] == true
  end)
end

local function inlines_have_note_refs(inlines)
  for _, it in ipairs(inlines or {}) do
    if it.type == "footnote_ref" then
      return true
    end
    if it.inlines and inlines_have_note_refs(it.inlines) then
      return true
    end
  end
  return false
end

local function block_has_note_refs(block)
  if inlines_have_note_refs(block.inlines) then
    return true
  end
  for _, cell in ipairs(block.headers or {}) do
    if inlines_have_note_refs(cell) then
      return true
    end
  end
  for _, row in ipairs(block.rows or {}) do
    for _, cell in ipairs(row) do
      if inlines_have_note_refs(cell) then
        return true
      end
    end
  end
  return false
end

local function max_numeric_attr(root, local_name, attr)
  local max_v = 0
  if root then
    for _, n in ipairs(xml.find_all(root, local_name)) do
      local v = tonumber(xml.get_attr(n, attr))
      if v and v > max_v then
        max_v = v
      end
    end
  end
  return max_v
end

-- Named colors accepted in color= attributes (anything else must be a hex value)
local COLOR_NAMES = {
  black = "000000", white = "FFFFFF", red = "FF0000", green = "00B050", blue = "0070C0",
  yellow = "FFFF00", orange = "FFC000", purple = "7030A0", gray = "808080", grey = "808080",
  darkred = "C00000", darkblue = "002060", lightblue = "00B0F0", lightgreen = "92D050",
}

local HIGHLIGHT_NAMES = {
  yellow = true, green = true, cyan = true, magenta = true, blue = true, red = true, darkBlue = true,
  darkCyan = true, darkGreen = true, darkMagenta = true, darkRed = true, darkYellow = true,
  darkGray = true, lightGray = true, black = true, white = true,
}

local function color_hex(value)
  if not value then
    return nil
  end
  local hex = value:gsub("^#", "")
  if hex:match("^%x%x%x%x%x%x$") then
    return hex:upper()
  end
  if hex:match("^%x%x%x$") then
    return (hex:sub(1, 1):rep(2) .. hex:sub(2, 2):rep(2) .. hex:sub(3, 3):rep(2)):upper()
  end
  return COLOR_NAMES[value:lower()]
end

-- <w:rPr> children for font / size (points) / color / highlight
run_attr_elements = function(fmt)
  local out = {}
  if fmt.font and fmt.font ~= "" then
    out[#out + 1] = xml.elem("w:rFonts", {
      ["w:ascii"] = fmt.font, ["w:hAnsi"] = fmt.font, ["w:cs"] = fmt.font, ["w:eastAsia"] = fmt.font,
    })
  end
  local size = tonumber(fmt.size)
  if size and size > 0 then
    local half_points = tostring(math.floor(size * 2 + 0.5))
    out[#out + 1] = xml.elem("w:sz", { ["w:val"] = half_points })
    out[#out + 1] = xml.elem("w:szCs", { ["w:val"] = half_points })
  end
  local color = color_hex(fmt.color)
  if color then
    out[#out + 1] = xml.elem("w:color", { ["w:val"] = color })
  end
  if fmt.highlight then
    local hl = fmt.highlight
    for name in pairs(HIGHLIGHT_NAMES) do
      if name:lower() == hl:lower() then
        hl = name
      end
    end
    if HIGHLIGHT_NAMES[hl] then
      out[#out + 1] = xml.elem("w:highlight", { ["w:val"] = hl })
    elseif color_hex(hl) then
      out[#out + 1] = xml.elem("w:shd", { ["w:val"] = "clear", ["w:color"] = "auto", ["w:fill"] = color_hex(hl) })
    end
  end
  return out
end

M.color_hex = color_hex

function M.ast_to_docx(doc, workspace, reference_archive, opts)
  opts = opts or {}
  local preserve = (opts.preserve_styles ~= false) and (reference_archive ~= nil)

  local archive = {
    files = {},
    order = {},
    -- Entries copied from the reference archive without being inflated (written back compressed)
    passthrough = {},
  }
  setmetatable(archive.files, {
    __index = function(t, name)
      local entry = archive.passthrough[name]
      if entry then
        local content = zip.materialize(entry)
        rawset(t, name, content)
        return content
      end
    end,
  })

  -- Copy one entry of the reference archive, without inflating it if nothing has read it yet
  local function copy_entry(name)
    if zip.is_unread(reference_archive, name) then
      archive.passthrough[name] = zip.unread_entry(reference_archive, name)
    else
      archive.files[name] = reference_archive.files[name]
    end
  end

  local preserved_sect_pr = nil
  local preserved_doc_attrs = nil
  local orig_doc_root = nil
  local preserved_doc_siblings = {}
  local orig_rels_by_id = {}
  local kept_rels = {}
  local max_rid_num = 0
  local header_footer_media = {}

  if preserve and reference_archive and reference_archive.files then
    archive.raw_entries = reference_archive.raw_entries
    local copied = {}
    local function copy_named(k)
      -- Copy all files except body-specific media that we will re-populate or retain
      if not copied[k] and not k:match("^word/media/") then
        copied[k] = true
        copy_entry(k)
        archive.order[#archive.order + 1] = k
      end
    end
    for _, name in ipairs(zip.names(reference_archive)) do
      copy_named(name)
    end
    for k, _ in pairs(reference_archive.files) do
      copy_named(k)
    end

    -- Inspect original word/document.xml for root attributes and <w:sectPr>
    if reference_archive.files["word/document.xml"] then
      local orig_doc = xml.parse(reference_archive.files["word/document.xml"])
      orig_doc_root = orig_doc
      if orig_doc then
        preserved_doc_attrs = vim.deepcopy(orig_doc.attrs or {})
        for _, c in ipairs(orig_doc.children or {}) do
          if c.tag and c.local_name ~= "body" then
            preserved_doc_siblings[#preserved_doc_siblings + 1] = c
          end
        end
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
            if zip.is_unread(reference_archive, mpath) then
              archive.passthrough[mpath] = zip.unread_entry(reference_archive, mpath)
              archive.order[#archive.order + 1] = mpath
            elseif rawget(reference_archive.files, mpath) then
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
      orig_rels_by_id[rel.id] = rel
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
  local styles_xml, heading_style_ids, body_style_id, style_name_to_id = ensure_required_styles(archive.files["word/styles.xml"])
  style_name_to_id = style_name_to_id or {}
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

  -- Run formatting inherited from the original paragraph an edited block replaces (font, size, color...)
  local base_rpr_children = nil

  local function merge_base_rpr(rpr_children)
    if not base_rpr_children then
      return rpr_children
    end
    local by_name = {}
    for _, c in ipairs(rpr_children) do
      by_name[c.local_name] = true
    end
    local merged = {}
    for _, c in ipairs(base_rpr_children) do
      if not by_name[c.local_name] then
        merged[#merged + 1] = vim.deepcopy(c)
      end
    end
    for _, c in ipairs(rpr_children) do
      merged[#merged + 1] = c
    end
    sort_rpr_children(merged)
    return merged
  end

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
        vim.list_extend(rpr_children, run_attr_elements(item))

        if not item.code then
          rpr_children = merge_base_rpr(rpr_children)
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
        elements[#elements + 1] = xml.elem("m:oMath", nil, latex_to_omml_nodes(item.text or ""))

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
        local isrc = item.src or ""
        if isrc:match("^charts/.*%.xml$") or isrc:match("^word/charts/.*%.xml$") then
          local chart_rel = isrc:gsub("^word/", "")
          local chart_zip = "word/" .. chart_rel
          if not archive.files[chart_zip] and reference_archive and reference_archive.files and reference_archive.files[chart_zip] then
            archive.files[chart_zip] = reference_archive.files[chart_zip]
            archive.order[#archive.order + 1] = chart_zip
          end
          local rid = add_rel(REL_CHART, chart_rel)
          local cx, cy = image.resolve_dimensions_emu(nil, item.width, item.height)
          pic_counter = pic_counter + 1
          local chart_drawing = build_chart_drawing_element(rid, pic_counter, item.alt or "", cx, cy)
          elements[#elements + 1] = xml.elem("w:r", nil, {
            chart_drawing,
          })
        else
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

  -- Reuse the original XML of blocks the user did not edit (mapping computed by preserve.lua)
  local psrc = preserve and opts.preserve_source or nil
  local pmap = (psrc and opts.preserve_map) or { exact = {}, inherit = {} }
  local orig_blocks = psrc and psrc.ast.blocks or {}
  local groups = psrc and psrc.ast._groups or {}
  local verbatim_group = {}
  local notes_rebuilt = false
  local reused_nodes = {}
  local glue_emitted = 0

  if psrc then
    -- Keep new bookmark and drawing ids clear of the ones in reused paragraphs
    bookmark_counter = max_numeric_attr(orig_doc_root, "bookmarkStart", "w:id")
    pic_counter = max_numeric_attr(orig_doc_root, "docPr", "id")

    local new_of_orig = {}
    for j, oi in pairs(pmap.exact) do
      new_of_orig[oi] = j
    end
    local candidate = {}
    for g, grp in ipairs(groups) do
      local j0 = new_of_orig[grp.first]
      local ok = j0 ~= nil
      for k = grp.first, grp.last do
        if not ok or new_of_orig[k] ~= j0 + (k - grp.first) then
          ok = false
          break
        end
      end
      candidate[g] = ok
    end

    -- footnotes.xml is rebuilt (and renumbered) as soon as any regenerated block references a note or a
    -- note's text changed; reused paragraphs pointing at the old note ids must then be regenerated too
    local new_notes = doc.footnotes or {}
    local old_notes = psrc.ast.footnotes or {}
    if #new_notes ~= #old_notes then
      notes_rebuilt = true
    else
      for k, fn in ipairs(new_notes) do
        local old = old_notes[k]
        if tostring(fn.id) ~= tostring(old.id)
          or markdown.render_inlines(fn.inlines or {}) ~= markdown.render_inlines(old.inlines or {}) then
          notes_rebuilt = true
          break
        end
      end
    end
    for j, block in ipairs(doc.blocks or {}) do
      local oi = pmap.exact[j]
      local g = oi and orig_blocks[oi]._group
      if not (g and candidate[g]) and block_has_note_refs(block) then
        notes_rebuilt = true
      end
    end
    for g, grp in ipairs(groups) do
      if candidate[g] then
        local has_notes = false
        for _, n in ipairs(grp.nodes) do
          if node_has_note_refs(n) then
            has_notes = true
          end
        end
        verbatim_group[g] = not (has_notes and notes_rebuilt)
      end
    end
  end

  -- Empty paragraphs, page and section breaks that sat in front of original groups
  local function emit_glue_through(g)
    while glue_emitted < math.min(g, #groups) do
      glue_emitted = glue_emitted + 1
      for _, n in ipairs(groups[glue_emitted].glue) do
        body_children[#body_children + 1] = n
        reused_nodes[#reused_nodes + 1] = n
      end
    end
  end

  -- Rewrite one table cell with new content, keeping its cell, paragraph and run properties
  local function rewrite_cell(tc, inlines)
    local first_p = xml.find_child(tc, "p")
    local ppr = inheritable_para_props(first_p)
    local saved_base = base_rpr_children
    base_rpr_children = inheritable_run_props(first_p)
    local p_children = compile_inlines(inlines)
    base_rpr_children = saved_base
    if ppr then
      table.insert(p_children, 1, ppr)
    end
    local kept = {}
    for _, c in ipairs(tc.children) do
      if c.tag and c.local_name == "tcPr" then
        kept[#kept + 1] = c
      end
    end
    kept[#kept + 1] = xml.elem("w:p", nil, p_children)
    tc.children = kept
  end

  -- An edited table with the same shape as the original: copy the original table and rewrite only the
  -- cells whose content changed. Returns nil when the shapes differ.
  local function patch_table(src, block)
    local tbl = src._pnode
    for _, c in ipairs(tbl.children or {}) do
      local ln = c.local_name
      if c.tag and ln ~= "tr" and ln ~= "tblPr" and ln ~= "tblGrid" and ln ~= "bookmarkStart" and ln ~= "bookmarkEnd" then
        return nil
      end
    end
    local new_rows = { block.headers or {} }
    for _, r in ipairs(block.rows or {}) do
      new_rows[#new_rows + 1] = r
    end
    local old_rows = { src.headers or {} }
    for _, r in ipairs(src.rows or {}) do
      old_rows[#old_rows + 1] = r
    end
    local copy = xml.clone(tbl)
    local trs = xml.find_children(copy, "tr")
    if #trs ~= #new_rows or #trs ~= #old_rows then
      return nil
    end
    local cell_opts = { in_table = true }
    local rewrites = {}
    for r, tr in ipairs(trs) do
      local col = 0
      for _, tc in ipairs(tr.children or {}) do
        if tc.tag and tc.local_name ~= "trPr" and tc.local_name ~= "tblPrEx" then
          if tc.local_name ~= "tc" then
            return nil
          end
          col = col + 1
          local first = col
          local tcpr = xml.find_child(tc, "tcPr")
          local gs = tcpr and xml.find_child(tcpr, "gridSpan")
          col = col + math.max(tonumber(gs and xml.get_attr(gs, "w:val")) or 1, 1) - 1
          for c = first, col do
            local new_md = markdown.render_inlines(new_rows[r][c] or {}, cell_opts)
            if new_md ~= markdown.render_inlines(old_rows[r][c] or {}, cell_opts) then
              if c ~= first then
                return nil
              end
              rewrites[#rewrites + 1] = { tc = tc, inlines = new_rows[r][c] or {} }
            end
          end
        end
      end
      if col ~= #old_rows[r] or #new_rows[r] > col then
        return nil
      end
    end
    for _, rw in ipairs(rewrites) do
      rewrite_cell(rw.tc, rw.inlines)
    end
    return copy
  end

  local generate_block_inner

  -- Paragraph properties from section attributes: style="Name" and align=left|center|right|justify
  local PPR_AFTER_JC = {
    textDirection = true, textAlignment = true, textboxTightWrap = true, outlineLvl = true, divId = true,
    cnfStyle = true, rPr = true, sectPr = true, pPrChange = true,
  }
  local ALIGN_VALUES = { left = "left", center = "center", centre = "center", right = "right", justify = "both", both = "both" }
  local function apply_block_ppr(node, battrs)
    local paras = node.local_name == "p" and { node } or xml.find_all(node, "p")
    for _, p in ipairs(paras) do
      local ppr = xml.find_child(p, "pPr")
      if not ppr then
        ppr = xml.elem("w:pPr")
        table.insert(p.children, 1, ppr)
      end
      if battrs.style then
        local sid = style_name_to_id[battrs.style:lower()] or battrs.style
        local ps = xml.find_child(ppr, "pStyle")
        if ps then
          xml.set_attr(ps, "w:val", sid)
        else
          table.insert(ppr.children, 1, xml.elem("w:pStyle", { ["w:val"] = sid }))
        end
      end
      local jc_val = battrs.align and ALIGN_VALUES[battrs.align:lower()]
      if jc_val then
        for k = #ppr.children, 1, -1 do
          if ppr.children[k].local_name == "jc" then
            table.remove(ppr.children, k)
          end
        end
        local at = #ppr.children + 1
        for k, c in ipairs(ppr.children) do
          if PPR_AFTER_JC[c.local_name] then
            at = k
            break
          end
        end
        table.insert(ppr.children, at, xml.elem("w:jc", { ["w:val"] = jc_val }))
      end
    end
  end

  local function set_pstyle(ppr, style_id)
    local ps = xml.find_child(ppr, "pStyle")
    if ps then
      xml.set_attr(ps, "w:val", style_id)
    else
      table.insert(ppr.children, 1, xml.elem("w:pStyle", { ["w:val"] = style_id }))
    end
  end

  -- Generate one block; `src` is the original block it replaces (if any), whose paragraph and run
  -- formatting carries over so an edited paragraph keeps its font, size, color, spacing and alignment
  local function generate_block(i, block, src)
    local btype = block.type
    local src_pnode = src and src.type == btype and src._pnode or nil

    if btype == "pagebreak" then
      body_children[#body_children + 1] = xml.elem("w:p", nil, {
        xml.elem("w:r", nil, { xml.elem("w:br", { ["w:type"] = "page" }) }),
      })
      return
    end

    -- ::: {...} section attributes: run defaults for everything inside, plus paragraph style/alignment
    local battrs = block.attrs and block.attrs.kv or nil
    if battrs then
      local run_defaults = {
        font = battrs.font,
        -- headings keep their own size; the section size is for body text
        size = btype ~= "heading" and battrs.size and tonumber((battrs.size:gsub("pt$", ""))) or nil,
        color = battrs.color,
        highlight = battrs.highlight,
      }
      local extra = run_attr_elements(run_defaults)
      if #extra > 0 then
        local replaced = {}
        for _, e in ipairs(extra) do
          replaced[e.local_name] = true
        end
        local merged = {}
        for _, c in ipairs(base_rpr_children or {}) do
          if not replaced[c.local_name] then
            merged[#merged + 1] = c
          end
        end
        vim.list_extend(merged, extra)
        base_rpr_children = merged
      end
      local first_new = #body_children + 1
      generate_block_inner(i, block, src, src_pnode)
      for k = first_new, #body_children do
        apply_block_ppr(body_children[k], battrs)
      end
      return
    end
    generate_block_inner(i, block, src, src_pnode)
  end

  generate_block_inner = function(i, block, src, src_pnode)
    local btype = block.type

    -- (a patched table keeps its original note references, which are only valid if notes are kept)
    if btype == "table" and src_pnode and not (notes_rebuilt and node_has_note_refs(src_pnode)) then
      local patched = patch_table(src, block)
      if patched then
        body_children[#body_children + 1] = patched
        reused_nodes[#reused_nodes + 1] = patched
        return
      end
    end
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
      local ppr = inheritable_para_props(src_pnode)
      if ppr then
        if (tonumber(src.level) or 1) ~= lvl then
          set_pstyle(ppr, style_id)
        end
      else
        ppr = xml.elem("w:pPr", nil, {
          xml.elem("w:pStyle", { ["w:val"] = style_id }),
        })
      end
      local p_children = {
        ppr,
        xml.elem("w:bookmarkStart", { ["w:id"] = bm_id, ["w:name"] = slug }),
      }
      for _, el in ipairs(compile_inlines(block.inlines)) do
        p_children[#p_children + 1] = el
      end
      p_children[#p_children + 1] = xml.elem("w:bookmarkEnd", { ["w:id"] = bm_id })
      body_children[#body_children + 1] = xml.elem("w:p", nil, p_children)

    elseif btype == "paragraph" then
      local p_children = compile_inlines(block.inlines)
      local ppr = inheritable_para_props(src_pnode)
      if not ppr and body_style_id then
        ppr = xml.elem("w:pPr", nil, { xml.elem("w:pStyle", { ["w:val"] = body_style_id }) })
      end
      if ppr then
        table.insert(p_children, 1, ppr)
      end
      body_children[#body_children + 1] = xml.elem("w:p", nil, p_children)

    elseif btype == "math_block" then
      body_children[#body_children + 1] = xml.elem("w:p", nil, {
        xml.elem("m:oMathPara", nil, {
          xml.elem("m:oMath", nil, latex_to_omml_nodes(block.text or "")),
        }),
      })

    elseif btype == "list_item" then
      local lvl = math.min(math.max(tonumber(block.level) or 0, 0), 8)
      local inherited_ppr = nil
      if src_pnode and (src.ordered or false) == (block.ordered or false) then
        inherited_ppr = inheritable_para_props(src_pnode)
        local num_pr = inherited_ppr and xml.find_child(inherited_ppr, "numPr")
        local ilvl = num_pr and xml.find_child(num_pr, "ilvl")
        if ilvl then
          xml.set_attr(ilvl, "w:val", tostring(lvl))
        end
      end
      local num_id
      if inherited_ppr then
        num_id = nil
      elseif block.ordered then
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
        inherited_ppr or xml.elem("w:pPr", nil, {
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
        inheritable_para_props(src_pnode) or xml.elem("w:pPr", nil, {
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

      local src_tblpr = src_pnode and xml.find_child(src_pnode, "tblPr")
      local tbl_children = {
        src_tblpr and vim.deepcopy(src_tblpr) or xml.elem("w:tblPr", nil, {
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

  for i, block in ipairs((doc and doc.blocks) or {}) do
    local oi = pmap.exact[i]
    local og = oi and orig_blocks[oi]._group
    if og and verbatim_group[og] then
      if oi == groups[og].first then
        emit_glue_through(og)
        for _, n in ipairs(groups[og].nodes) do
          body_children[#body_children + 1] = n
          reused_nodes[#reused_nodes + 1] = n
        end
      end
      active_ordered_num_id = nil
    else
      local src_i = oi or pmap.inherit[i]
      local src = src_i and orig_blocks[src_i] or nil
      if src and src._group then
        emit_glue_through(src._group)
      end
      local style_i = not src and pmap.style_from and pmap.style_from[i] or nil
      if style_i then
        -- Look only: a copy without the neighbour's section break
        src = vim.tbl_extend("force", {}, orig_blocks[style_i])
        src._pnode = src._pnode and xml.clone(src._pnode)
        local ppr = src._pnode and xml.find_child(src._pnode, "pPr")
        if ppr then
          for k = #ppr.children, 1, -1 do
            if ppr.children[k].local_name == "sectPr" then
              table.remove(ppr.children, k)
            end
          end
        end
      end
      base_rpr_children = src and src.type == block.type and inheritable_run_props(src._pnode) or nil
      generate_block(i, block, src)
      base_rpr_children = nil
    end
  end
  if psrc then
    emit_glue_through(#groups)
    for _, n in ipairs(psrc.ast._trailing_glue or {}) do
      body_children[#body_children + 1] = n
      reused_nodes[#reused_nodes + 1] = n
    end

    -- Relationships (images, links, charts, embedded objects) used by reused XML keep their original ids
    local rids = {}
    for _, n in ipairs(reused_nodes) do
      collect_rel_ids(n, rids)
    end
    local have = {}
    for _, r in ipairs(kept_rels) do
      have[r.id] = true
    end
    for rid in pairs(rids) do
      local rel = orig_rels_by_id[rid]
      if rel and not have[rid] then
        have[rid] = true
        kept_rels[#kept_rels + 1] = rel
        if rel.type == REL_IMAGE and rel.target_mode ~= "External" then
          local mpath = normalize_zip_path("word", rel.target)
          local present = rawget(archive.files, mpath) ~= nil or archive.passthrough[mpath] ~= nil
          if not present and (zip.is_unread(reference_archive, mpath) or reference_archive.files[mpath]) then
            copy_entry(mpath)
            archive.order[#archive.order + 1] = mpath
            local ext = (mpath:match("%.(%w+)$") or ""):lower()
            if ext ~= "" then
              media_exts[ext] = image.content_type_for_ext(ext)
            end
          end
        end
      end
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

  local doc_children = {}
  for _, c in ipairs(preserved_doc_siblings) do
    doc_children[#doc_children + 1] = c
  end
  doc_children[#doc_children + 1] = xml.elem("w:body", nil, body_children)
  local doc_elem = xml.elem("w:document", doc_attrs, doc_children)
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

  -- Ensure _rels/.rels exists and its officeDocument target points to word/document.xml
  if not archive.files["_rels/.rels"] then
    archive.files["_rels/.rels"] = xml.serialize(
      xml.elem("Relationships", { ["xmlns"] = NS_PKG_REL }, {
        xml.elem("Relationship", { ["Id"] = "rId1", ["Type"] = REL_OFFICE_DOC, ["Target"] = "word/document.xml" }),
        xml.elem("Relationship", { ["Id"] = "rId2", ["Type"] = REL_CORE_PROPS, ["Target"] = "docProps/core.xml" }),
        xml.elem("Relationship", { ["Id"] = "rId3", ["Type"] = REL_EXT_PROPS, ["Target"] = "docProps/app.xml" }),
      })
    )
  else
    local rroot = xml.parse(archive.files["_rels/.rels"])
    if rroot then
      local rmod = false
      for _, rel_node in ipairs(xml.find_children(rroot, "Relationship")) do
        if xml.get_attr(rel_node, "Type") == REL_OFFICE_DOC then
          local t = xml.get_attr(rel_node, "Target") or ""
          if t ~= "word/document.xml" and t ~= "/word/document.xml" then
            rel_node.attrs["Target"] = "word/document.xml"
            rmod = true
          end
        end
      end
      if rmod then
        archive.files["_rels/.rels"] = xml.serialize(rroot)
      end
    end
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

-- Write every word/media/* entry to workspace/media/ (inflating entries skipped at load time)
function M.extract_media(archive, workspace)
  for _, name in ipairs(zip.names(archive)) do
    local media_sub = name:match("^word/(media/.+)$")
    if media_sub and vim.fn.filereadable(workspace .. "/" .. media_sub) == 0 then
      local content = archive.files[name]
      if content then
        write_binary_file(workspace .. "/" .. media_sub, content)
      end
    end
  end
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
    for _, name in ipairs(zip.names(archive)) do
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

