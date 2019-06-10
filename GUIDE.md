## Guide to basic functionality
### Opening the file

You can open ODS file (OpenDocument Spreadsheet) like this
````ruby
workbook = Rspreadsheet.open('./test.ods')
workbook = Rspreadsheet.open('./test.fods') # gem supports flast OpenDocument format
````
and access its first sheet like this
````ruby
sheet = workbook.worksheet(1)
````
### Accessing cells

you can get and set contents of cells using "verbatim" syntax like
````ruby
sheet.row(5).cell(4).value 
sheet.row(5).cell(4).value = 10
````
or using "brief" syntax like
````ruby
sheet[5,4]
sheet[5,4] = 10
````

You can mix these two at will, for example like this
````ruby
row = sheet.row(5)
row[4] = 10
````

### Working with images
sheet.insert_image_to('10.21mm','15mm','image.png')
i = sheet.images.first
i.move_to('100mm','99.98mm')

### Saving 
The file needs to be saved after doing changes. 
````ruby
workbook.save
workbook.save('new_filename.ods')    # changes filename and saves
workbook.save(any_io_object)         # file can be saved to any IO like object as well
workbook.to_io                       # coverts it to IO object which can be used to 
anotherIO.write(workbook.to_io.read) # send file over internet without saving it first
````

### Creating fresh new file
You may name the spreadsheet on creation or at first save.

````ruby
workbook = Rspreadsheet.new
workbook.save('./filename.ods')      # filename nust be provided at least on first save
workbook2 = Rspreadsheet.new('./filename2.fods', format: :flat)
workbook2.save                       
```

If you want to use the fods flat format, you must create it as such.

### Date and Time
OpenDocument and ruby have different models of date, time and datetime. Ruby containg three different objects. Time and DateTime cover all cases, Date covers dates only. OpenDocument distinguishes two groups - time of a day (time) and everything else (date). To simplify things a little we return cell values containg time of day as Time object and cell values containg datetime of date as DateTime. I am aware that this is very arbitrary choice, but it is very practical. This way and to some extend the types of values from OpenDocument are preserved when read from files, beeing acted upon and written back to spreadshhet.

### Merged cells 
Even when a cell spans more rows and collumns, you must access by coordinates of its topleft corner. In fact, the "hidden" cells under are still there and their content is usually preserved.


## Examples

  * [basic functionality](https://gist.github.com/gorn/42e33d086d9b4fda10ec) 
  * [extended examples](https://gist.github.com/gorn/b432e6a69e82628349e6) of lots of alternative syntax

## Conventions
  * **all indexes are 1-based**. This applies to rows, cells cordinates, and all array like structures like list od  worksheets etc. Spreadsheet world is 1-based, ruby is 0-based so I had to make a decision. I intend to make an global option for this, but in early stage I need to keep things simple.
  * with numeric coordinates row always comes before col as in  (row,col)
  * with alphanumerical col always comes before row as in F12
  * Shorter syntax worksheet[x,y] returns value, longer syntax worksheet.cell(x,y) return cell objects. This allows to work conviniently with values using short syntax and access the cell object if needed (to access formatting for example).
  * Singular and plural like sheet/sheets, row/rows, cell/cells can be used intergangebly.

