-- Maps the blocks of an edited Markdown buffer back to the blocks of the document as it was loaded,
-- so the writer can reuse the original WordprocessingML for everything the user did not touch
-- (keeping direct formatting such as font sizes, colors, spacing and alignment) and let edited
-- blocks inherit the formatting of the paragraph they replace.

local markdown = require("nvim-docx.markdown")

local M = {}

-- Snapshot taken when a document is loaded: the Markdown shown to the user, the line range each
-- original block rendered to, and the AST (whose blocks carry _pnode/_group from ooxml.docx_to_ast).
function M.make_source(lines, ranges, ast)
  if not ast or not ast._groups then
    return nil
  end
  return {
    lines = lines,
    ranges = ranges,
    ast = ast,
  }
end

-- Returns { exact = {...}, inherit = {...}, style_from = {...} }, each mapping new_idx -> orig_idx
--   exact:      the new block's Markdown is byte-identical to an original block's Markdown
--   inherit:    the new block is an edited version of (or replacement for) that original block
--   style_from: a brand-new paragraph that should look like the nearest original paragraph
function M.map_blocks(source, new_lines, new_ranges, new_blocks)
  local result = { exact = {}, inherit = {}, style_from = {} }
  if not source then
    return result
  end

  local old_lines = source.lines
  local orig_blocks = source.ast.blocks
  local old_ranges = source.ranges

  local old_text = table.concat(old_lines, "\n") .. "\n"
  local new_text = table.concat(new_lines, "\n") .. "\n"
  local diff = vim.text and vim.text.diff or vim.diff
  local hunks = diff(old_text, new_text, { result_type = "indices", algorithm = "histogram" }) or {}

  -- new line -> old line for every line outside a changed hunk
  local new2old = {}
  local a, b = 1, 1
  local function copy_until(next_a, next_b)
    while a < next_a and b < next_b do
      new2old[b] = a
      a, b = a + 1, b + 1
    end
  end
  for _, h in ipairs(hunks) do
    local sa, ca, sb, cb = h[1], h[2], h[3], h[4]
    -- For pure insertions/deletions xdiff reports the line *before* the gap
    local ha = (ca == 0) and (sa + 1) or sa
    local hb = (cb == 0) and (sb + 1) or sb
    copy_until(ha, hb)
    a, b = ha + ca, hb + cb
  end
  copy_until(#old_lines + 1, #new_lines + 1)

  -- old line -> original block index
  local old_line_block = {}
  for i, r in pairs(old_ranges) do
    for l = r[1], r[2] do
      old_line_block[l] = i
    end
  end

  -- Which original blocks have a new block exactly equal to them
  for j, block in ipairs(new_blocks) do
    local r = new_ranges[j]
    if r then
      local first_old = new2old[r[1]]
      local i = first_old and old_line_block[first_old]
      local orig = i and orig_blocks[i]
      if orig and orig.type == block.type and old_ranges[i][1] == first_old
        and markdown.block_attrs_sig(orig.attrs) == markdown.block_attrs_sig(block.attrs)
        and (old_ranges[i][2] - old_ranges[i][1]) == (r[2] - r[1]) then
        local all_same = true
        for l = r[1], r[2] do
          if new2old[l] ~= first_old + (l - r[1]) then
            all_same = false
            break
          end
        end
        if all_same then
          result.exact[j] = i
        end
      end
    end
  end

  -- Edited blocks: inherit from the original block that shares an unchanged line with it
  local used = {}
  for _, i in pairs(result.exact) do
    used[i] = true
  end
  for j, block in ipairs(new_blocks) do
    local r = new_ranges[j]
    if r and not result.exact[j] then
      for l = r[1], r[2] do
        local ol = new2old[l]
        local i = ol and old_line_block[ol]
        if i and not used[i] and orig_blocks[i].type == block.type then
          result.inherit[j] = i
          used[i] = true
          break
        end
      end
    end
  end

  -- Fully rewritten blocks: pair the remaining new and original blocks inside each changed hunk in order
  for _, h in ipairs(hunks) do
    local sa, ca, sb, cb = h[1], h[2], h[3], h[4]
    if ca > 0 and cb > 0 then
      local olds, news = {}, {}
      local seen = {}
      for l = sa, sa + ca - 1 do
        local i = old_line_block[l]
        if i and not used[i] and not seen[i] then
          seen[i] = true
          olds[#olds + 1] = i
        end
      end
      for j, r in pairs(new_ranges) do
        if r[1] >= sb and r[1] <= sb + cb - 1 and not result.exact[j] and not result.inherit[j] then
          news[#news + 1] = j
        end
      end
      table.sort(news)
      for k = 1, math.min(#olds, #news) do
        local i, j = olds[k], news[k]
        if orig_blocks[i].type == new_blocks[j].type then
          result.inherit[j] = i
          used[i] = true
        end
      end
    end
  end

  -- New list items next to an original list item join that list (so numbering continues)
  for j, block in ipairs(new_blocks) do
    if block.type == "list_item" and not result.exact[j] and not result.inherit[j] then
      for _, nj in ipairs({ j - 1, j + 1 }) do
        local ni = result.exact[nj] or result.inherit[nj]
        local nb = new_blocks[nj]
        local ob = ni and orig_blocks[ni]
        if ob and nb and nb.type == "list_item" and ob.type == "list_item"
          and (nb.ordered or false) == (block.ordered or false) and (nb.level or 0) == (block.level or 0) then
          result.inherit[j] = ni
          break
        end
      end
    end
  end

  -- Brand-new paragraphs take their look (font, size, spacing) from the nearest original paragraph,
  -- so typing a new paragraph does not fall back to the template's Normal style
  local styled = { paragraph = true, blockquote = true }
  for j, block in ipairs(new_blocks) do
    if styled[block.type] and not result.exact[j] and not result.inherit[j] then
      for dist = 1, #new_blocks do
        local found = nil
        for _, nj in ipairs({ j - dist, j + dist }) do
          local ni = result.exact[nj] or result.inherit[nj]
          if ni and orig_blocks[ni].type == block.type then
            found = ni
            break
          end
        end
        if found then
          result.style_from[j] = found
          break
        end
        if j - dist < 1 and j + dist > #new_blocks then
          break
        end
      end
    end
  end

  return result
end

return M
