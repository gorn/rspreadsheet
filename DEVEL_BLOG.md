## Conventions
  * with numeric coordinates row always comes before col as in  [row,col]
  * with alphanumerical col always comes before row as in F12

## Ideas/wishlist

  * Shouldn't some of the alternative syntax (like sheet.cells[x,y]) return cell objects, while normal sheet[cell] would just return value?
  * We should determine whether sheet[0,0] should be base type like String or Rspreadsheet::Cell.
 
## Developing this gem

### How to test

  * by running <code>guard</code> in terminal you will get tested any changes as soon as they are summitted


### Automated utilities
 
  * [travis-ci](https://travis-ci.org/gorn/rspreadsheet/jobs/25375065) provides automated testing
  * [github](https://github.com/gorn/rspreadsheet) hosts the repository where you can get the code
  * [coverals](https://coveralls.io/r/gorn/rspreadsheet) checks how well source is covered by tests
