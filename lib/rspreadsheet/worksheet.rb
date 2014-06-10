require 'rspreadsheet/row'
require 'forwardable'

module Rspreadsheet
class Worksheet
  attr_accessor :name
  extend Forwardable
  def_delegators :@worksheetcells, :nonemptycells

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
# this object is result of sheet.cells
class WorksheetCells
  def initialize
    @cells = Hash.new do |hash, coords| 
      # we create empty cell and place it to hash, we do not have to check whether there is a cell in XML already, because it would be in hash as well
      hash[coords]=Cell.new(coords[0],coords[1])
      # TODO: create XML empty node here or upon save?
    end
  end
  def [](r,c)
    cells_object(r,c)
  end
  def nonemptycells
    @cells.values
  end
  
  ### internal
  def cells_object(r,c)
    @cells[[r,c]]
  end
  def initialize_cell(r,c,source_node)
    @cells[[r,c]]=Cell.new(r,c,source_node)
  end
end

end
