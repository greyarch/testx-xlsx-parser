xlsx = require './lib/xlsx'

xlsx.parse('/tmp/test.xlsx', 'Sheet1').then (script) -> console.log JSON.stringify(script, null, 2)
