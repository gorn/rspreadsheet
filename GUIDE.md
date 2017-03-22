## Guide to basic functionality
### Opening the file

You can open ODS file (OpenDocument Spreadsheet) like this
````ruby
@workbook = Rspreadsheet.open('./test.ods')
````
and access its first sheet like this
````ruby
@sheet = @workbook.worksheet(1)
````
### Accessing cells

you can get and set contents of cells using "verbatim" syntax like
````ruby
@sheet.row(5).cell(4).value 
@sheet.row(5).cell(4).value = 10
````
or using "brief" syntax like
````ruby
@sheet[5,4]
@sheet[5,4] = 10
````

You can mix these two at will, for example like this
````ruby
@row = @sheet.row(5)
@row[4] = 10
````

### Working with images
@sheet.insert_image_to('10.21mm','15mm','image.png')
i = @sheet.images.first
i.move_to('100mm','99.98mm')

### Saving 
The file needs to be saved after doing changes. 
````ruby
@workbook.save
@workbook.save('new_filename.ods')   # changes filename and saves
@workbook.save(any_io_object)        # file can be saved to any IO like object as well
````

## Examples

  * [basic functionality](https://gist.github.com/gorn/42e33d086d9b4fda10ec) 
  * [extended examples](https://gist.github.com/gorn/b432e6a69e82628349e6) of lots of alternative syntax

## Conventions
  * **all indexes are 1-based**. This applies to rows, cells cordinates, and all array like structures like list od  worksheets etc. Spreadsheet world is 1-based, ruby is 0-based so I had to make a decision. I intend to make an global option for this, but in early stage I need to keep things simple.
  * with numeric coordinates row always comes before col as in  (row,col)
  * with alphanumerical col always comes before row as in F12
  * Shorter syntax worksheet[x,y] returns value, longer syntax worksheet.cell(x,y) return cell objects. This allows to work conviniently with values using short syntax and access the cell object if needed (to access formatting for example).
  * Note: currently plural and singular like sheet/sheets, row/rows, cell/cells can be used intergangebly, but there is a (discussion)[https://github.com/gorn/rspreadsheet/issues/10] about this and in future versions we might use singular style for methods and plural for array style.

