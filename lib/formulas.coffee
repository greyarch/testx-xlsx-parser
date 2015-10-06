formulajs = require 'formulajs'
_ = require 'lodash'
moment = require 'moment'

text = formulajs.TEXT
today = -> moment().diff(moment([1900, 0, 1]), 'days')

module.exports =
  _.extend formulajs,
    TODAY: today
    NOW: today
    TEXT: (n, fmt) ->
      dateFormat = fmt.toUpperCase().replace /J/g, 'Y' # Dutch to English date format
      fDate = moment([1900, 0, 1]).add(n, 'days').format(dateFormat)
      if moment(fDate, dateFormat).isValid()
        fDate
      else
        text n, fmt
    CELL: ->
      console.warn 'Warning: "CELL" Excel function is not implemented!'
    INDIRECT: ->
      console.warn 'Warning: "INDIRECT" Excel function is not implemented!'
