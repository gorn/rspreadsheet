## Examples of advanced syntax

Examples of more advanced syntax follows

```ruby
require 'rspreadsheet'

book = Rspreadsheet.new
sheet = book.create_worksheet 'Top icecreams'

sheet[1,1] = 'My top 5'
p sheet[1,1].class           # => String
p sheet[1,1]                 # => "My top 5"

# These are all the same values - alternative syntax
p sheet.A1
p sheet[1,1]
p sheet.cells(0,0).value
p sheet.rows(0).cells(0).value   

# How to inspect/manipulate the Cell object
sheet.cells(1,1)            # => Rspreadsheet::Cell
sheet.cells(1,1).format
sheet.cells(1,1).format.size = 15
sheet.cells(1,1).format.weight = bold
p sheet.cells(1,1).format.bold?  # => true

# There are the same assigmenents
sheet.A1 = value
sheet[1,1]= value
sheet.cells(1,1).value = value

p sheet.A1.class          # => Rspreadsheet::Cell

# build the top five list
(1..5).each { |i| sheet[i,1] = i }
sheet.columns(1).format.bold = true
sheet.cells[2,1..5] = ['Vanilla', 'Pistacia', 'Chocolate', 'Annanas', 'Strawbery']

sheet.columns(1).cells(1).format.color = :red

book.save

```
## Conventions
  * with numeric coordinates row always comes before col as in  (row,col)
  * with alphanumerical col always comes before row as in F12
  * Shorter syntax worksheet[x,y] returns value, longer syntax worksheet.cells(x,y) return cell objects. This allows to work conviniently with values using short syntax and access the cell object if needed (for formatting for example).

