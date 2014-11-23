# Changelog
## 0.2.4 version
  * new accessors like sheet.cells('A3') or sheet['A3']='hi' ro even sheet['A','3'] etc.

## 0.2.x versions (x<4)

  * complete rewrite without xmlnodes caching
  * adding formats to cells
  * refactoring common parts of code from row and cell to XMLTiedItem and XMLTiedArray objects
  * increasing travis coverage
  * cell outside used range now returns "UncreatedRow/Cell" which is detached upon value assignement. This solves the problem that spreadsheed is virtually infinite in Spreadsheet word, but in reality it has clear boundaries.
  * Accessors for nonempty/defined cells.
  
## 0.1.x versions
  * initial take on syntax and all fuctions, it is usable now
