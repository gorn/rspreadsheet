# rspreadsheet

rspreadsheet - manipulating spreadsheets with Ruby. Read, modify, write or create new OpenDocument Spreadsheet files from ruby code.

## Examples
  
```ruby
require 'rspreadsheet'

book = Rspreadsheet.open('./existing_file.ods')
sheet = book.worksheets 'Icecream list'
total = 0

sheet.rows[3..20].each do |row|
  puts 'Icecream name: " + row[2]
  puts 'Icecream ingredients: ' + row[3]
  puts "I ate this " + row[4] + ' times'
  total += row[4]
end

sheet[21,3] = 'Total:'
sheet[21,4] = total

sheet.D21 = 'Total'
sheet.D21.format.bold = true

book.save

```

Some examples of alternative syntax

```ruby
require 'rspreadsheet'

book = Rspreadsheet::Workbook.new
sheet book.create_worksheet 'Top icecreams'

sheet[0,0] = 'My top 5'
sheet[0,0].format.size = 15
sheet[0,0].format.weight = bold
p sheet[0,0].format.bold  # => true

# These are all the same cells
p sheet.A1
p sheet.row(0).cell(0)    
p sheet.rows[0][0]  
p sheet.rows[0].cells[0]
p sheet.cells[0,0]
p sheet.cell(0,0)

p sheet.A1.class          # => Rspreadsheet::Cell

# build the top ten list
(1..5).each { |i| sheet[i,0] = i }
sheet.columns[0].format.bold = true
sheet.cells[1,1..5] = ['Vanilla', 'Pistacia', 'Chocolate', 'Annanas', 'Strawbery']

sheet.columns[1][1..3].format.color = :red

book.save

```




## Installing

To install the gem run the following

    gem install rspreadsheet
