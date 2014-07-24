require 'rspreadsheet/row'
require 'forwardable'

module Rspreadsheet

module Tools
  ## converts cell adress like 'F12' to pair od integers [row,col]
  def self.convert_cell_address(*coords)
    if coords.length == 1
      coords.match(/^([A-Z]{1,3})(\d{1,8})$/)
      colname = $~[1]
      rowname = $~[2]
    elsif coords.length == 2
      colname = coords[0]
      rowname = coords[1]
    else
      raise 'Wrong number of arguments'
    end
      
    colname=colname.rjust(3,'@')
    col = (colname[-1].ord-65)+(colname[-2].ord-64)*26+(colname[-3].ord-64)*26*26
    row = rowname.to_i-1
    return [row,col]
  end      
end

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
  def rows
    WorksheetRows.new(self)
  end
  def method_missing method_name, *args, &block
    if method_name.to_s.match(/^([A-Z]{1,3})(\d{1,8})(=?)$/)
      row,col = Tools.convert_cell_address($~[1],$~[2])
      assignchar = $~[3]
      if assignchar == '='
        self.cells[row,col].value = args.first
      else
        self.cells[row,col].value
      end
    else
      super
    end
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

# this allows the sheet.rows[r] syntax
# this object is result of sheet.rows
class WorksheetRows
  def initialize(aworkbook)
    @workbook = aworkbook
    @spredsheetrows=Array.new()
  end
  def [] rowi
    @spredsheetrows[rowi] ||= Row.new(@workbook,rowi)
  end
end

end
