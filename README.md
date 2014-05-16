# rspreadsheet

rspreadsheet - manipulating spreadsheets with Ruby. Read, modify, write or create new OpenDocument Spreadsheet files from ruby code.

## Status

This project is its brainstorming phase. Nothing is implemented yet, the documentation now serves as a list of intentions. Please submit issues and/or fork the repository if you have more ideas, wishes, etc ...


## Examples
  
```ruby
require 'rspreadsheet'

book = Rspreadsheet.open('./existing_file.ods')
sheet = book.worksheets 'Icecream list'
total = 0

sheet.rows[3..20].each do |row|
  puts 'Icecream name: ' + row[2]
  puts 'Icecream ingredients: ' + row[3]
  puts "I ate this " + row[4] + ' times'
  total += row[4]
end

sheet[21,3] = 'Total:'
sheet[21,4] = total

sheet.rows[21].format.bold = true

book.save

```

This is the basic functionality. However rspreadsheet should allows lots of alternative syntax, like described in (GUIDE.md)



## Installing

To install the gem run the following

    gem install rspreadsheet
