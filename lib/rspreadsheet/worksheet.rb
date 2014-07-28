require 'rspreadsheet/row'
require 'rspreadsheet/tools'
require 'forwardable'

module Rspreadsheet

class Worksheet
  attr_accessor :name, :xmlnode
  extend Forwardable
  def_delegators :nonemptycells

  def initialize(xmlnode_or_sheet_name)
    # set up the @xmlnode according to parameter
    case xmlnode_or_sheet_name
      when LibXML::XML::Node
        @xmlnode = xmlnode_or_sheet_name
      when String
        @xmlnode = LibXML::XML::Node.new('table:table')
        @xmlnode['table:name'] = xmlnode_or_sheet_name
      else raise 'Provide name or xml node to create a Worksheet object'
    end
      
    ## initialize rows
    @spredsheetrows=RowArray.new(@xmlnode)
  end
  def cells(r,c)
    rows(r).cells(c)
  end
  def nonemptycells
    @cells.values
  end
  def rows(rowi)
    @spredsheetrows.get_row(rowi)
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
