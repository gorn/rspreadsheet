require 'rspreadsheet/row'
require 'rspreadsheet/tools'
# require 'forwardable'

module Rspreadsheet 

class Worksheet
  include XMLTiedArray
  attr_accessor :name, :xmlnode
  def subitem_xml_options; {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'} end

  def initialize(xmlnode_or_sheet_name)
    @itemcache=Hash.new    
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
  end
  
  def rowxmlnode(rowi)
    find_my_subnode_respect_repeated(rowi, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
  
  def row_nonempty_cells_col_indexes(rowi)
    arowxmlnode = rowxmlnode(rowi)
    if arowxmlnode.nil?
      []
    else
      find_nonempty_subnode_indexes(arowxmlnode, {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'})
    end
  end
  
  def cellxmlnode(rowi,coli)
    arowxmlnode = rowxmlnode(rowi)
    if arowxmlnode.nil?
      nil
    else
      rors(rowi).find_my_subnode_respect_repeated(coli, {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'})
    end
  end
 
  def first_unused_row_index
    find_first_unused_index_respect_repeated({:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
  
  def insert_row_above(arowi)
    insert_subitem_before(arowi)
  end
  
  def insert_cell_before(arowi,acoli)
    detach_row_in_xml(arowi)
    rows(arowi).insert_subitem_before(acoli)
  end
  
  def detach_row_in_xml(rowi)
    return detach_my_subnode_respect_repeated(rowi, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
  
  def nonemptycells
    used_rows_range.collect{ |rowi| rows(rowi).nonemptycells }.flatten
  end
  
  # rozšíření XMLTiedArray
  def rows(rowi); subitem(rowi) end
  def prepare_subitem(rowi); Row.new(self,rowi) end
  def rowcache; @itemcache end
  
  ## syntactic sugar follows
  def [](r,c)
    cells(r,c).andand.value
  end
  def []=(r,c,avalue)
    cells(r,c).andand.value=avalue
  end
  def cells(r,c)
    rows(r).andand.cells(c)
  end
  # allows syntax like sheet.F15
  def method_missing method_name, *args, &block
    if method_name.to_s.match(/^([A-Z]{1,3})(\d{1,8})(=?)$/)
      row,col = Rspreadsheet::Tools.convert_cell_address_to_coordinates($~[1],$~[2])
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
    1..self.first_unused_row_index-1
  end
end

end
