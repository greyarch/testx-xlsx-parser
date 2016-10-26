xlsx = require '../lib/xlsx'

script = xlsx.parse('test/test.xlsx', 'Sheet1')
console.log JSON.stringify(script, null, 2)
