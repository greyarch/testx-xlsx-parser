q = require 'q'
xlsx = require 'xlsx'

exports.parse = (xlsFile, sheet) ->
  deferred = q.defer()
  wb = xlsx.readFile xlsFile
  rows = getRows wb.Sheets[sheet]
  steps = for row, i in rows
    getKeyword(rows, i)
  deferred.resolve
    steps: (s for s in steps when s)
  deferred.promise

getRows = (sheet) ->
  r = /([a-zA-Z]+)(\d+)/i
  rows = []
  for key, value of sheet
    if key[0] != '!'
      [match, col, row] = r.exec key
      row = parseInt row
      col = col.toUpperCase()
      unless rows[row] then rows[row] = []
      rows[row][col] = value
  rows

getKeyword = (rows, i) ->
  row = rows[i]
  prevRow = rows[i - 1]
  if row?['A']?.w
    [match, name, ignore, comment] = /([^\[|\]]*)(\[(.*)\])?/.exec row['A'].w
    keyword =
      name: name.trim()
      meta:
        Row: i
        "Full name": row['A'].w
        Comment: comment?.trim() || ''
      arguments: {}
    cols = Object.keys(rows[i - 1]).sort (l, r) -> l > r
    for col in cols
      keyword.arguments[prevRow[col]?.w] = row[col]?.w || ''
    keyword
  else
    null
