## Examples

  * [basic functionality](https://gist.github.com/gorn/42e33d086d9b4fda10ec) 
  * [extended examples](https://gist.github.com/gorn/b432e6a69e82628349e6) of lots of alternative syntax

## Conventions
  * **all indexes are 1-based**. This applies to rows, cells cordinates, and all array like structures like list od  worksheets etc. Spreadsheet world is 1-based, ruby is 0-based do I had to make a decision. I intend to make an global option for this, but in early stage I need to keep things simple. 
  * with numeric coordinates row always comes before col as in  (row,col)
  * with alphanumerical col always comes before row as in F12
  * Shorter syntax worksheet[x,y] returns value, longer syntax worksheet.cells(x,y) return cell objects. This allows to work conviniently with values using short syntax and access the cell object if needed (to access formatting for example).

