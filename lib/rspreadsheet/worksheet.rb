require 'rspreadsheet/row'
require 'rspreadsheet/tools'
# require 'forwardable'

module Rspreadsheet

class Worksheet
  attr_accessor :name, :xmlnode
#   extend Forwardable
#   def_delegators :nonemptycells

  def initialize(xmlnode_or_sheet_name)
    # set up the @xmlnode according to parameter
    case xmlnode_or_sheet_name
      when LibXML::XML::Node
        @xmlnode = xmlnode_or_sheet_name
      when String
        @xmlnode = LibXML::XML::Node.new('table')
        ns = LibXML::XML::Namespace.new(@xmlnode, 'table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0')
        @xmlnode .namespaces.namespace = ns
        @xmlnode['table:name'] = xmlnode_or_sheet_name
      else raise 'Provide name or xml node to create a Worksheet object'
    end
      
    ## initialize rows
    @spredsheetrows=RowArray.new(self,@xmlnode)
  end
  def cells(r,c)
    rows(r).andand.cells(c)
  end
  def nonemptycells
    used_rows_range.collect{ |rowi| rows(rowi) }.collect { |row| row.nonemptycells }.flatten
  end
  def rows(rowi)
    @spredsheetrows.get_row(rowi)
  end
  ## syntactic sugar follows
  def [](r,c)
    cells(r,c).andand.value
  end
  def []=(r,c,avalue)
    cells(r,c).andand.value=avalue
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
  def used_rows_range
    1..@spredsheetrows.first_unused_row_index-1
  end
end

end
