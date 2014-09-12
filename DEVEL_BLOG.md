See [GUIDE.md](GUIDE.md#conventions) for syntax conventions.

## Ideas/wishlist

  * What should be returned when asking for row/cell outside used range? Currently it creates new row/cell on fly, but maybe it should only return nil, so the user needs to insert apropriate rows/cells himself before using them. However it spoils little bit syntax like spreadsheet.rows(5).cells(3) because if rows returns nil, that would fail and ugly constructs like spreadsheet.andand.rows(5).andand.cells(3) would be needed. The only way this could be acoided is by using something like "UncreatedRow" object. This concern falls to category "clash of worlds" like the indexing issue (0 vs 1 based) and many others.
  * In future inntroduce syntax like ``sheet.range('C3:E4')`` for mass operations. Also maybe ``sheet.cells('C3')`` or ``sheet.cells(3, 'C')`` etc.
  * Trying to make row Enumerable - perhaps skipping empty or undefined cells.
  * Accessors for nonempty/defined cells.
  * Maybe insted two syntaxes for accessing cell, we make both of them do the same and return Proxy object which would act either as value or cell.
  * Allow any of these:
    * ``book['Spring 2014']`` as alternative to ``book.worksheets('Spring 2014')`` ?
    * ``sheet.cells.F13`` as alternative to ``sheet.cells[14,5]`` ?

Guiding ideas
  * xml document is always synchronized with the data. So the save is trivial.
  * no duplication of data. Objects like RowArray should containg minimum information. This one exists solely to speed up cell search. Taken to extream it is questionable, whether we need such objects at all, it might be possible to always work with xml directly.

    
## Developing this gem

### Automated testing

  * ``budele exec guard`` will get tested any changes as soon as they are summitted
  * ``rake spec`` runs all test manually

### Automated utilities
 
  * [travis-ci](https://travis-ci.org/gorn/rspreadsheet) provides automated testing
  * [github](https://github.com/gorn/rspreadsheet) hosts the repository where you can get the code
  * [coverals](https://coveralls.io/r/gorn/rspreadsheet) checks how well source is covered by tests

### Local manual testing and releasing (to github released, ).

    gem build rspreadsheet.gemspec
    sudo gem install rspreadsheet-x.y.z.gem
    gem push rspreadsheet-x.y.z.gem

alternative way using ``rake`` command - release is more automatic

  * ``rake build`` - builds the gem to pkg directory. 
  * ``sudo rake install`` - If this fails with "mkmf.rb can't find header files for ruby at /usr/lib/ruby/include/ruby.h" you may want to ``sudo aptitude install ruby-dev``
  * Now can locally and manually use/test the gem. This should not be replacement for automated testing. 
  * ``rake release`` - creates a version tag in git and pushes the code to github + Rubygems. After this is succesfull the new version appears as release in Github and RubyGems.



### Naming conventions

  * create_xxx - creates object and inserts it where necessary
  * prepare_xxx - create object, but does not insert it anywhere
