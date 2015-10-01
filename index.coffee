xlsx = require './xlsx'

module.exports =
  parse: xlsx.parse

xlsx.parse('/tmp/test.xlsx', 'Sheet1').then (res) -> console.log JSON.stringify res, null, 2
