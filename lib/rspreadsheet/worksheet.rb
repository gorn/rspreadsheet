require 'rspreadsheet/row'

module Rspreadsheet
class Worksheet
  attr_accessor :name

  def initialize
    @worksheetcells=WorksheetCells.new
  end

  def cells
    @worksheetcells
  end

  def [](r,c)
    cells[r,c].value
  end

  def []=(r,c,avalue)
    cells[r,c].value=avalue
  end
  
end

# this allows the sheet.cells[r,c] syntax
# this onject is result od sheet.cells
class WorksheetCells
  def initialize
    @cells = Hash.new{ |hash, coords| hash[coords]=Cell.new() }
  end
  
  def [](r,c)
    cells_object(r,c)
  end
  
  # internal
  def cells_object(r,c)
    @cells[[r,c]]
  end
  
#   def row(r)
#     @rows[r] || @rows[r] = Row.new(self,r)  # the association to the @rows should be written on creation, otherwise more desinchronized copies may exist. TODO:How amd when to handle writing to xml
#   end  
end

end
