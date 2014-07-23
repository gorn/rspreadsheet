require('rspreadsheet/cell')

# Currently this is only syntax sugar for cells and contains no functionality

module Rspreadsheet

class Row
  def initialize(workbook,rowi)
    @rowi = rowi
    @workbook = workbook
    @rowcells = RowCells.new(workbook,rowi)
  end
  def cells
    @rowcells
  end
end

# this allows the row.cells[c] syntax
# this object is result of row.cells
class RowCells
  def initialize(workbook,rowi)
    @rowi = rowi
    @workbook = workbook
  end
  def [] coli
    @workbook.cells[@rowi,coli]
  end
end

end
