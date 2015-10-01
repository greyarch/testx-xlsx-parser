q = require 'q'
xlsx = require 'xlsx'

formulas = require './formulas'

exports.parse = (xlsFile, sheet) ->
  deferred = q.defer()
  wb = xlsx.readFile xlsFile
  calcWb wb.Sheets
  scriptSheet = wb.Sheets[sheet]
  rows = getRows scriptSheet
  steps = for row, i in rows
    getKeyword(rows, i)
  deferred.resolve
    steps: (s for s in steps when s)
  deferred.promise

calcWb = (sheets) ->
  resolveRef = (sheetName, ref) ->
    cell = sheets[sheetName][ref]
    if cell.f and not cell.calculated
      cell.v = cell.calculated = calc sheetName, ref
    else
      cell.v

  calc = (sheetName, ref) ->
    cell = sheets[sheetName][ref]
    ff = cell.f.replace(/[^!A-Z]([A-Z]+[0-9]+)|(^[A-Z]+[0-9]+)/g, 'resolveRef(sheetName, "$1$2")')
    ff = ff.replace(/([A-Za-z0-9]+)!([A-Z]+[0-9]+)/g, 'resolveRef("$1", "$2")')
    ff = ff.replace(/([A-Z]+)\(/g, 'formulas.$1(')
    eval ff

  calcSheet = (name, sheet) ->
    for cellRef, cellVal of sheet
      if cellVal.f
        resolveRef name, cellRef

  for shName, sh of sheets
    calcSheet shName, sh

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
  if row?['A']?.v
    [match, name, ignore, comment] = /([^\[|\]]*)(\[(.*)\])?/.exec row['A'].v
    keyword =
      name: name.trim()
      meta:
        Row: i
        "Full name": row['A'].v
        Comment: comment?.trim() || ''
      arguments: {}
    cols = Object.keys(rows[i - 1]).sort (l, r) -> l > r
    for col in cols
      keyword.arguments[prevRow[col]?.v] = row[col]?.v || ''
    keyword
  else
    null
