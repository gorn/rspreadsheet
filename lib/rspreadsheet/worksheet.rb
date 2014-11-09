require 'rspreadsheet/row'
require 'rspreadsheet/tools'
# require 'forwardable'

module Rspreadsheet 

class Worksheet
  attr_accessor :name, :xmlnode
#   extend Forwardable
#   def_delegators :nonemptycells

  def initialize(xmlnode_or_sheet_name)
    @rowcache=[]
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
    find_subnode_respect_repeated(@xmlnode, rowi, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
 
  def rowrange(rowi)
    find_subnode_range_respect_repeated(@xmlnode, rowi, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
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
      find_subnode_respect_repeated(arowxmlnode, coli, {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'})
    end
  end
 
  def cellrange(coli)
    find_subnode_range_respect_repeated(@xmlnode, rowi, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
  
  def first_unused_row_index
    find_first_unused_index_respect_repeated(xmlnode, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
  
  def detach_row_in_xml(rowi)
    return detach_subnode_respect_repeated(xmlnode, rowi, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
  end
  def detach_cell_in_xml(rowi,coli)
    rownode = detach_row_in_xml(rowi)
    return detach_subnode_respect_repeated(rownode, coli, {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'})
  end
  
  def detach_subnode_respect_repeated(axmlnode,aindex, options)
    index = 0
    axmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      oldindex = index
      index = index+repeated
      if index>= aindex  # found the node, now do the detachement
        ranges = [oldindex+1..aindex-1,aindex..aindex,aindex+1..index].reject {|range| range.size<1}
        ranges.each do |range|
          newnode = node.copy(true)
          Tools.set_ns_attribute(newnode,'table',options[:xml_repeated_attribute],range.size,1)
          node.prev = newnode
        end
        node.remove!
        return find_subnode_respect_repeated(axmlnode, aindex, options)
      end
    end
    # add outbound xmlnode
    [index+1..aindex-1,aindex..aindex].reject {|range| range.size<1}.each do |range|
      node = LibXML::XML::Node.new(options[:xml_items_node_name],nil, Tools.get_namespace('table'))
      Tools.set_ns_attribute(node,'table',options[:xml_repeated_attribute],range.size, 1)
      axmlnode << node
    end  
    find_subnode_respect_repeated(axmlnode, aindex, options)
  end

  def find_subnode_respect_repeated(axmlnode, aindex, options)
    index = 0
    axmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      index = index+repeated
      return node if index>= aindex
    end
    return nil
  end
  def find_subnode_range_respect_repeated(axmlnode, aindex, options)
    index = 0
    axmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      if index+repeated >= aindex
        return (index+1..index+repeated)
      else
        index = index+repeated
      end
    end
    return (index+1..Float::INFINITY)
  end
  
  def find_nonempty_subnode_indexes(axmlnode, options)
    index = 0
    result = []
    axmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      index = index + repeated
      if !(node.content.nil? or node.content.empty? or node.content =='') and (repeated==1)
        result << index
      end
    end
    return result
  end
  def find_first_unused_index_respect_repeated(axmlnode, options)
    index = 0
    axmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      index = index+repeated
    end
    return index+1
  end
  
  def cells(r,c)
    rows(r).andand.cells(c)
  end
  def nonemptycells
    used_rows_range.collect{ |rowi| rows(rowi) }.collect { |row| row.nonemptycells }.flatten
  end
  def rows(rowi)
    @rowcache[rowi] ||= Row.new(self,rowi) unless rowi<=0
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
