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
  formulaUtils = require './utils'

  resolveRef = (sheetName, ref) ->
    [targetSheet, targetRef] = [sheetName, ref]
    if crossSheetRef = ref.match /^'?([^\[\]\*\?\:\/\\']+)'?!([A-Z]+[0-9]+)/
      [targetSheet, targetRef] = crossSheetRef[1..2]
    cell = sheets[targetSheet]?[targetRef]
    if cell?.f and not cell?.calculated
      cell.v = cell.calculated = calc targetSheet, targetRef
    else
      cell?.v or ''

  calc = (sheetName, ref) ->
    cell = sheets[sheetName]?[ref]
    eval formulaUtils.format(cell.f.replace(/&/g, '+'), sheetName)

  calcSheet = (name, sheet) ->
    for cellRef, cellVal of sheet
      if cellVal.f
        resolveRef name, cellRef
      else
        cellVal.v = cellVal.w

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
    if rows[i - 1] # are there any arguments at all?
      cols = Object.keys(rows[i - 1]).sort (l, r) -> l > r
      for col in cols
        keyword.arguments[prevRow[col]?.v] = row[col]?.v || ''
    keyword
  else
    null
