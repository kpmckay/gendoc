## gendoc
Replaces keys in Word (.docx) document with values from a CSV file.

## Dependencies
lxml: https://lxml.de/installation.html

python-docx: https://python-docx.readthedocs.io/en/latest/user/install.html

## Basic Idea
If a Word document needs data from another source, it is sometimes easier to manage that data separately in a spreadsheet. This script replaces placeholders in a Word document with values from a spreadsheet table (exported as CSV). The spreadsheet need only have "key" and "value" columns. Other columns can be added to make management easier (e.g. a description of the value). At the moment, the script insists on keys in the Word document being encapsulated in tables. It is a "todo" item to make it more flexible. There is probably some way to do this in Word, but it seemed easy enough to do in Python. 
