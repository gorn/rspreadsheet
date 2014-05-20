Examples of more advanced syntax follows:

```ruby
require 'rspreadsheet'

book = Rspreadsheet.new
sheet = book.create_worksheet 'Top icecreams'

sheet[0,0] = 'My top 5'
p sheet[0,0].class           # => String
p sheet[0,0]                 # => "My top 5"

# These are all the same values - alternative syntax
p sheet.A1
p sheet[0,0]
p sheet.cells[0,0].value
p sheet.rows[0].cells[0].value   

# How to inspect/manipulate the Cell object
sheet.cells[0,0]            # => Rspreadsheet::Cell
sheet.cells[0,0].format
sheet.cells[0,0].format.size = 15
sheet.cells[0,0].format.weight = bold
p sheet.cells[0,0].format.bold?  # => true

# There are the same assigmenents
sheet.A1 = value
sheet[0,0]= value
sheet.cells[0,0].value = value

p sheet.A1.class          # => Rspreadsheet::Cell

# build the top ten list
(1..5).each { |i| sheet[i,0] = i }
sheet.columns[0].format.bold = true
sheet.cells[1,1..5] = ['Vanilla', 'Pistacia', 'Chocolate', 'Annanas', 'Strawbery']

sheet.columns[1][1..3].format.color = :red

book.save

```
