testx-xlsx-parser
=====

Simple XLSX file parser for use with the testx library. Converts a test script in a XLSX sheet to testx test script (JSON).

## API
This library exposes only one method **parse** that takes 2 arguments *fileName* and *sheetName* and returns a promise that resolves to the JSON representation of the test script on that sheet. Example:
  
```
  xlsx = require('testx-xlsx-parser')
  xlsx.parse('xls-files/test.xlsx', 'Sheet1')
```
