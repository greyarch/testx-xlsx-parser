formulajs = require 'formulajs'
_ = require 'lodash'
moment = require 'moment'

text = formulajs.TEXT
module.exports =
  _.extend formulajs,
    TEXT: (n, fmt) ->
      if isNaN(n.toString())
        moment(n).format(fmt.toUpperCase())
      else
        text n, fmt
    CELL: ->
      console.warn 'Warning: "CELL" Excel function is not implemented!'
    INDIRECT: ->
      console.warn 'Warning: "INDIRECT" Excel function is not implemented!'
