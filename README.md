testx-pdf-keywords
=====

A library that extends testx with keywords for testing PDF files. This library is packaged as a npm package

## How does it work
From the directory of the art code install the package as follows:
```sh
npm install testx-pdf-keywords --save
```

After installing the package add the keywords to your protractor config file as follows:

```
testx.addKeywords(require('testx-pdf-keywords'))
```

## Keywords

| Keyword                | Argument name | Argument value  | Description | Supports repeating arguments |
| ---------------------- | ------------- | --------------- |------------ | ---------------------------- |
| check in pdf      |               |                 | check that the expected regex matches the text in the PDF file |  |
|                        | file           | full path to the pdf file; one of **file**, **url** or **link** has to be specified || No |
|                        | url           | URL of the pdf file; one of **file**, **url** or **link** has to be specified || No |
|                        | link           | link to the pdf file; one of **file**, **url** or **link** has to be specified || No |
|                        | timeout        | timeout in milliseconds to wait for the link to the PDF to appear on the screen; ignored if the keyword is used with **file** or **url**;  parameters; defaults to 5000 if not present || No |
|                        | expect1(2, 3...) | expected regular expression to match the text in the pdf || Yes |
| check not in pdf  |               |                 | same as **check in pdf**, but checks that the specified regex DOES NOT match |  |
