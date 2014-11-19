# Changelog
## 0.2.x versions

  * complete rewrite without xmlnodes caching
  * adding formats to cells
  * refactoring common parts of code from row and cell to XMLTiedItem and XMLTiedArray objects
  * increasing travis coverage
  * cell outside used range now returns "UncreatedRow/Cell" which is detached upon value assignement. This solves the problem that spreadsheed is virtually infinite in Spreadsheet word, but in reality it has clear boundaries.
  
## 0.1.x versions
  * initial take on syntax and all fuctions, it is usable now
