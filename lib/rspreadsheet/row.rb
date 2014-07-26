require('rspreadsheet/cell')

# Currently this is only syntax sugar for cells and contains no functionality

module Rspreadsheet

class Row
  def initialize(workbook,rowi)
    @rowi = rowi
    @workbook = workbook
  end
  def cells(coli)
    @workbook.cells(@rowi,coli)
  end
end

end
