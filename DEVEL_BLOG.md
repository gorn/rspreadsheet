See [GUIDE.md](GUIDE.md#conventions) for syntax conventions.

## Ideas/wishlist

  * Trying to make row Enumerable - perhaps skipping empty or undefined cells.
  * Accessors for nonempty/defined cells.
  * Maybe insted two syntaxes for accessing cell, we make both of them do the same and return Proxy object which would act either as value or cell.
  * Allow any of these? 
    * book['Spring 2014'] in place of book.worksheets('Spring 2014')
 
## Developing this gem

### How to test

  * by running <code>guard</code> in terminal you will get tested any changes as soon as they are summitted

### Automated utilities
 
  * [travis-ci](https://travis-ci.org/gorn/rspreadsheet/jobs/25375065) provides automated testing
  * [github](https://github.com/gorn/rspreadsheet) hosts the repository where you can get the code
  * [coverals](https://coveralls.io/r/gorn/rspreadsheet) checks how well source is covered by tests

### Notes

Submitting to rubygems.org

    gem build rspreadsheet.gemspec
    gem push rspreadsheet-x.y.z.gem

