formulajs = require 'formulajs'
_ = require 'lodash'
moment = require 'moment'

text = formulajs.TEXT
today = -> moment().diff(moment([1900, 0, 1]), 'days') + 2

module.exports =
  _.extend formulajs,
    TODAY: today
    NOW: today
    CELL: ->
      console.warn 'Warning: "CELL" Excel function is not implemented!'
    INDIRECT: ->
      console.warn 'Warning: "INDIRECT" Excel function is not implemented!'
