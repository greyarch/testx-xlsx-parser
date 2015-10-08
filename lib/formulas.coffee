formulajs = require 'formulajs'
_ = require 'lodash'
moment = require 'moment'
ssf = require('xlsx').SSF

today = -> moment().diff(moment([1900, 0, 1]), 'days') + 2

module.exports =
  _.extend formulajs,
    TEXT: (n, fmt) -> ssf.format fmt, n
    TODAY: today
    NOW: today
    CELL: ->
      console.warn 'Warning: "CELL" Excel function is not implemented!'
    INDIRECT: ->
      console.warn 'Warning: "INDIRECT" Excel function is not implemented!'
