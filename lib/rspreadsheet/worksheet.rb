require 'rspreadsheet/row'

module Rspreadsheet
class Worksheet
  attr_accessor :name

  def initialize(source_node=nil)
    @source_node = source_node
    @worksheetcells=WorksheetCells.new
    rowi = 0
    unless @source_node.nil?
      @source_node.elements.select{ |node| node.name == 'table-row'}.each do |row_source_node|
        coli = 0
        row_source_node.elements.select{ |node| node.name == 'table-cell'}.each do |cell_source_node|
          cells.initialize_cell(rowi,coli,cell_source_node)
          coli += 1
        end
        rowi += 1
      end    
    end
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
  
  def initialize_cell(r,c,source_node)
    @cells[[r,c]]=Cell.new(source_node)
  end
#   def row(r)
#     @rows[r] || @rows[r] = Row.new(self,r)  # the association to the @rows should be written on creation, otherwise more desinchronized copies may exist. TODO:How amd when to handle writing to xml
#   end  
end

end
