

require 'rspreadsheet'
book = Rspreadsheet.open './time.ods'
sheet = book.worksheets(1)
# cell = sheet.A1 # exception
cell = sheet.cells('A1') # works
cell.to_s # exception