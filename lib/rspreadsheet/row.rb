require('rspreadsheet/cell')
include Forwardable

# Currently this is only syntax sugar for cells and contains no functionality

module Rspreadsheet

class RowArray
  def initialize(aworksheet_node)
    @worksheet_node = aworksheet_node

    # initialize @rowgroups from @worksheet_node
    @rowgroups = []
    unless @worksheet_node.nil?
      @worksheet_node.elements.select{|node| node.name == 'table-row'}.each do |row_source_node|
        @rowgroups << prepare_row_group(row_source_node) # it is in @worksheet_node so suffices to add object to @rowgroups
      end
    end
  end
  def prepare_row_group(size_or_xmlnode)  # appends new RowGroup at the end
    # reading params
    if size_or_xmlnode.kind_of? LibXML::XML::Node
      size = (size_or_xmlnode['number-cols-repeated'] || 1).to_i
      node = size_or_xmlnode
    elsif size_or_xmlnode.to_i>0
      size = size_or_xmlnode.to_i
      node = nil
    else
      return nil
    end
    index = first_unused_row_index
    
    # construct result
    RowGroup.new(self,index..index+size-1,node)
#     .normalize
  end
  def add_row_group(size_or_xmlnode)
    result = prepare_row_group(size_or_xmlnode)
    @rowgroups << result
    @worksheet_node << result.xmlnode
    result
  end
  def get_row_group(rowi)
    @rowgroups.find{ |rowgroup| rowgroup.range.cover?(rowi) }
  end
  def get_row(rowi)
    rg = get_row_group(rowi).andand.normalize
    case rg
      when SingleRow then rg
      when RowGroup then MemberOfRowGroup.new(rowi, rg)
      when nil then get_out_of_bound_row_group(rowi)
      else raise
    end
  end
  # prolonges the RowArray to cantain rowi and returns it
  def get_out_of_bound_row_group(rowi)
    fill_row_group_size = rowi-first_unused_row_index
    add_row_group(fill_row_group_size) if fill_row_group_size>0
    add_row_group(1)
  end
  def first_unused_row_index
    if @rowgroups.empty? 
      1
    else
      @rowgroups.last.range.end+1
    end
  end
  # This detaches row rowi from the group and perhaps splits the RowGroup
  # into two pieces. This makes the row individually editable.
  def detach(rowi)
    index = get_row_group_index(rowi)
    row_group = @rowgroups[index]
    range = row_group.range

    # prepare new components
    replaceby = []
    replaceby << RowGroup.new(self,range.begin..rowi-1)
    replaceby << (result = SingleRow.new(self,rowi))
    replaceby << RowGroup.new(self,rowi+1..range.end)
    
    # put original range somewhere in replaceby and shorten it
    if rowi>range.begin
      replaceby[0] = row_group
      row_group.range = range.begin..rowi-1
    else
      replaceby[2] = row_group
      row_group.range = rowi+1..range.end
    end
    
    # normalize and delete empty parts
    replaceby = replaceby.map(&:normalize).compact
    
    # do the replacement in xml
    marker = LibXML::XML::Node.new('temporarymarker')
    row_group.xmlnode.next = marker
    row_group.xmlnode.remove!
    replaceby.each{ |rg| 
      puts rg.inspect
      marker.prev = rg.xmlnode
    } 
    marker.remove!
    
    # do the replacement in array
    @rowgroups[index..index]=replaceby
    result
  end
  
 private
  def get_row_group_index(rowi)
    @rowgroups.find_index{ |rowgroup| rowgroup.range.cover?(rowi) }
  end
end

class Row
  @readonly = :unknown
  # ? @rowindex 
end

class RowWithXMLNode < Row
  attr_accessor :xmlnode
  def style_name=(value)
    @xmlnode['style-name'] = value
  end
  def cells(coli)
    cellnode = nil
    while true 
      curr_coli=1
      cellnode = @xmlnode.elements.select{|n| n.name=='table-cell'}.find do |el|
        curr_coli += (el['number-cols-repeated'] || 1).to_i
        curr_coli > coli
      end
      unless cellnode.nil? 
        return Cell.new(self,coli,cellnode)
      else
        add_cell
      end
    end
  end
  def add_cell(repeated=1)
    cell = Cell.new(self,first_unused_column_index)
    cell.xmlnode[':number-columns-repeated'] = repeated.to_s if repeated>1
    @xmlnode << cell.xmlnode
    cell
  end
  def used_range
    fu = first_unused_column_index
    (fu>1) ? 1..fu : nil
  end
  def first_unused_column_index
    1 + @xmlnode.elements.select{|n| n.name=='table-cell'}.reduce(0) do |sum, el|
      (el.attributes.get_attribute_ns('urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'number-cols-repeated') || 1).to_i + sum
    end
  end
end

class RowGroup < RowWithXMLNode
  @readonly = :yes_always
  attr_accessor :range, :parent_array, :xmlnode
  def initialize(aparent_array,arange,axmlnode=nil)
    @parent_array = aparent_array
    @range = arange
    if axmlnode.nil?
      ns = Tools.get_namespace('table')
      axmlnode = LibXML::XML::Node.new('table-row')
      axmlnode.namespaces.namespace = ns
      axmlnode['number-rows-repeated']=range.size.to_s
    end
    @xmlnode = axmlnode
  end
  # returns SingleRow if size of range is 1 and nil if it is 0 or less
  def normalize
    case range.size
      when 2..Float::INFINITY then self
      when 1 then SingleRow.new_from_rowgroup(self)
      else nil
    end
  end
  def repeated;  range.size   end
  def repeated?; range.size>1 end

end

class SingleRow < RowWithXMLNode
  @readonly = :no
  attr_accessor :xmlnode
  # index  Integer
  def initialize(aparent_array,aindex,axmlnode=nil)
    @parent_array = aparent_array
    @index = aindex
    if axmlnode.nil?
      axmlnode = LibXML::XML::Node.new('table:table-row')
    end
    @xmlnode = axmlnode
  end
  def self.new_from_rowgroup(rg)
    anode = rg.xmlnode
    anode['table:number-rows-repeated'] = '1'
    
    SingleRow.new(rg.parent_array,rg.range.begin,anode)
  end
  def normalize; self end
  def repeated?; false end
  def repeated; 1 end
  def range; (@index..@index) end
  def detach; true end
  def row; @index end
end

## there are not data in this object, they are taken from RowGroup, but this is only readonly
class MemberOfRowGroup < Row
  @readonly = :yes_but_detachable
  extend Forwardable
  delegate [:repeated?, :repeated, :xmlnode, :parent_array] => :@row_group
  attr_accessor :row_group # for dubugging

  # @index  Integer
  # @row_group   RepeatedRow
  def initialize(aindex,arow_group)
    @index = aindex.to_i
    @row_group = arow_group
    raise 'Wrong parameter given' unless @row_group.is_a? RowGroup
  end
  def detach  # detaches MemberOfRowGroup from its RowGroup perhaps splitting RowGroup
    @row_group.parent_array.detach(@index)
  end
  def row
    @index
  end
end

end

## was in Row
#   def initialize(workbook,rowi)
#     @rowi = rowi
#     @workbook = workbook
#   end
#   def cells(coli)
#     @workbook.cells(@rowi,coli)
#   end
