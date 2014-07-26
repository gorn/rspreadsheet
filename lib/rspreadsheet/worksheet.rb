require 'rspreadsheet/row'
require 'rspreadsheet/tools'
require 'forwardable'

module Rspreadsheet

class Worksheet
  attr_accessor :name
  extend Forwardable
  def_delegators :nonemptycells

  def initialize(source_node=nil)
    @source_node = source_node
    ## initialize cells
    @cells = Hash.new do |hash, coords|
      # we create empty cell and place it to hash, we do not have to check whether there is a cell in XML already, because it would be in hash as well
      hash[coords]=Cell.new(coords[0],coords[1])
      # TODO: create XML empty node here or upon save?
    end
    rowi = 1
    unless @source_node.nil?
      @source_node.elements.select{ |node| node.name == 'table-row'}.each do |row_source_node|
        coli = 1
        row_source_node.elements.select{ |node| node.name == 'table-cell'}.each do |cell_source_node|
          initialize_cell(rowi,coli,cell_source_node)
          coli += 1
        end
        rowi += 1
      end
    end
    ## initialize rows
    @spredsheetrows=Array.new()
  end
  def initialize_cell(r,c,source_node)
    @cells[[r,c]]=Cell.new(r,c,source_node)
  end
  def cells(r,c)
    @cells[[r,c]]
  end
  def nonemptycells
    @cells.values
  end
  def rows(rowi)
    @spredsheetrows[rowi] ||= Row.new(self,rowi)
  end
  ## syntactic sugar follows
  def [](r,c)
    cells(r,c).value
  end
  def []=(r,c,avalue)
    cells(r,c).value=avalue
  end
  # allows syntax like sheet.F15
  def method_missing method_name, *args, &block
    if method_name.to_s.match(/^([A-Z]{1,3})(\d{1,8})(=?)$/)
      row,col = Rspreadsheet::Tools.convert_cell_address($~[1],$~[2])
      assignchar = $~[3]
      if assignchar == '='
        self.cells(row,col).value = args.first
      else
        self.cells(row,col).value
      end
    else
      super
    end
  end
end

end
