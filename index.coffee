xlsx = require './lib/xlsx'

module.exports =
  parse: xlsx.parse

xlsx.parse('/tmp/test.xlsx', 'Sheet1').then console.log 
