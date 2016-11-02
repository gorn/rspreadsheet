See [GUIDE.md](GUIDE.md#conventions) for syntax conventions.

## Installation troubleshooting

After you have cloned the source run `bundle` command in gem directory to install needed tools. If there are any errors try to

  1. Make sure that necessary tools are installed. In debian based distros you may try this  `apt-get install make gcc`. In other distros, use their native way to install tools. 
  2. If you have installed `ruby-libxml` than you may want to comment out the line containing libxml-ruby dependency in rspreadseet.gemspec file before running bundle.
  3. Comment out `guard` and `guard-rspec` lines from gemfile if you use ruby version less than 2.2.1.

## Ideas/wishlist

  * In future inntroduce syntax like `sheet.range('C3:E4')` for mass operations.
  * Trying to make row Enumerable - perhaps skipping empty or undefined cells.
  * Maybe insted two syntaxes for accessing cell, we make both of them do the same and return Proxy object which would act either as value or cell.
  * Document that there is a little distinction betwean RSpreadsheet and RSpreadsheet::Workbook. The former delegates everythink to the other.
  * allow `book.worskheets.first` syntax
  * allow `sheet.cells.sum { |cell| cell.value }
  * allow `sheet.rows(1).cells.each {}  iterate through nonempty cells ??? or all of them until last used
  * `sheet.rows(1).cells` returns list of cells objects and `sheet.rows(1).cellvalues` return array of values with nils when empty
  * implement to_csv
  * longterm plan - go through other used libraries and try to find out whose syntax could be adopted, so this library is drop in replacement (possibly with some config options) for them
  * iterative generation like this
 
 ```ruby
RSpreadsheet.generate('pricelist.ods') do 
     row 'icecream name', 'price'
     { 'Vanilla' => 2, 'Banana' => 3, 'Strawbery' => 2.7 }.each do |icecream, price|
       row icecream, price
       row '2x 'icecream, price * 1.8
     end
     skip_row
     row 'Menu made by rspreadsheet', :format => {:font => {:size => 5}}
     move_to A5
     cell 'Have a nice day!'
 end
```

* possible ideas from [Google Spreadsheet API](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet) (issue #9). For now just raw list of what could be used
   * appendRaw(rowContents)
   * concept of active sheets with methods like - deleteActiveSheet, getActiveSheet, duplicateActiveSheet(), getActiveCell()
  * `@sheet.delete_column(2), @sheet.delete_columns(from, howmany), @sheet.delete_row(2), ...`
  * GS API uses `@workbook.deletesheet(@sheet)` but it is nor very ruby way. What about `@worksheet.delete` or `@workbook.delete_sheet(2)` as syntactic shugar to `@workbook.worksheets(2).delete`
  * getAs(contentType) - Return the data inside this object as a blob converted to the specified content type. Maybe something like `@workbook.worksheets(2).to_pdf` or `@workbook.worksheets(2).to_blob('application/pdf')`. Is there anything like Blob in ruby?
  * getDataRange() - Returns a Range corresponding to the dimensions in which data is present
  * setActiveRange(range),  setActiveSelection(range) - ? what is the difference
  * @sheet. getFrozenColumns(),  @sheet.getFrozenColumns() 
  * getLastColumn(),  getLastRow()  
  * getRange(a1Notation), getRangeByName(name), range.get_values
  * getRowHeight(rowPosition)
  * hideColumn(column), unhideColumn(column), insertColumnAfter(afterPosition), insertColumnsAfter(afterPosition, howMany), insertImage(blob, column, row)
  * moveActiveSheet(pos) 
  * sort(columnPosition,ascending=true) - Sorts a sheet by column. 
  * Range#breakApart() - Break any multi-column cells in the range into individual cells again.
  * Range#canEdit()
  * Range#copyTo(destination)
  * Range#merge() - Merges the cells in the range together into a single block.
* @book.sheet_names - Array of names of sheets.


* array returned by rows of cells can have predefined dummy object as defaults for out of range indexes


##Guiding ideas
  * xml document is always synchronized with the data. So the save is trivial.
  * no duplication of data. Objects like RowArray should containg minimum information. This one exists solely to speed up cell search. Taken to extream it is questionable, whether we need such objects at all, it might be possible to always work with xml directly.
  * all cells and rows only server as proxy. they hold index and worksheet pointer and everytime read or write is done, the xml is newly searched. until there is a xmlnode caching we have no problem
  * all cells returned for particular coordinates are **always the same object**. This way we have no problem to synchronize changes.
    
## Known Issues
  * currently there is a confusing syntax @worksheet.rows(1).cells[5] which returns a cell object in SIXT column, which is not intended. It is side effecto of Row#cells returning an array
    
## Developing this gem

### Automated testing

  * ``bundle exec guard`` will get tested any changes as soon as they are summitted
  * ``rake spec`` runs all test manually

### Automated utilities
 
  * [travis-ci](https://travis-ci.org/gorn/rspreadsheet) provides automated testing
  * [github](https://github.com/gorn/rspreadsheet) hosts the repository where you can get the code
  * [coverals](https://coveralls.io/r/gorn/rspreadsheet) checks how well source is covered by tests

### Local testing and releasing (to github and rubygems).

1. Make changes
2. Test if all tests pass (run `bundle exec guard` to test automatically). If not go to 1.
3. Build and install locally using script `./reinstall_local_gem.sh` (see [details](reinstall_local_gem.sh))
4. Now can locally and manually use/test the gem. This should not be replacement for automated testing. If you make some changes, go to 1.
5. When happy, increment the version number and `git add .; git commit -am'commit message'; git push`
6. ``rake release`` - creates a version tag in git and pushes the code to github + Rubygems. After this is succesfull the new version appears as release in Github and RubyGems.

gem alternativa to points 3-6

    gem build rspreadsheet.gemspec              -\   These two lines together are in install.sh
    sudo gem install rspreadsheet-x.y.z.gem     -/   which should be invoked from parent directory
    gem push rspreadsheet-x.y.z.gem             releases the gem, do not forgot to update version in rspreadsheet.gemspec before doing this

### Naming conventions

  * create_xxx - creates object and inserts it where necessary
  * prepare_xxx - create object, but does not insert it anywhere
