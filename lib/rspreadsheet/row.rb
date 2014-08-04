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
      size = (size_or_xmlnode['number-rows-repeated'] || 1).to_i
      node = size_or_xmlnode
    elsif size_or_xmlnode.to_i>0
      size = size_or_xmlnode.to_i
      node = nil
    else
      return nil
    end
    index = first_unused_row_index
    
    # construct result
    RowGroup.new(self,index..index+size-1,node).normalize
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
      when nil
       if rowi>0 then UninitializedEmptyRow.new(self,rowi) else nil end
      else raise
    end
  end
  # prolonges the RowArray to cantain rowi and returns it
  def detach_of_bound_row_group(rowi)
    fill_row_group_size = rowi-first_unused_row_index
    if fill_row_group_size>0
      add_row_group(fill_row_group_size) 
    end
    add_row_group(1)
    return get_row(rowi)
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
  def self.empty_row_node
    LibXML::XML::Node.new('table-row',nil, Tools.get_namespace('table'))
  end
end

class RowWithXMLNode < Row
  attr_accessor :xmlnode
  def style_name=(value); Tools.set_ns_attribute(@xmlnode,'table','style-name',value)  end
  def cells(coli)
    coli = coli.to_i
    return nil if coli.to_i<=0
    Cell.new(self,coli,cellnodes(coli))
  end
  def nonemptycells
    nonemptycellsindexes.collect{ |index| cells(index) }
  end
  def nonemptycellsindexes
    used_col_range.to_a.select do |coli|
      cellnode = cellnodes(coli)
      !(cellnode.content.nil? or cellnode.content.empty? or cellnode.content =='') or
      !cellnode.attributes.to_a.reject{ |attr| attr.name == 'number-columns-repeated'}.empty?
    end
  end
  def used_col_range
    1..first_unused_column_index-1
  end
  def cellnodes(coli)
    cellnode = nil
    while true 
      curr_coli=1
      cellnode = @xmlnode.elements.select{|n| n.name=='table-cell'}.find do |el|
        curr_coli += (Tools.get_ns_attribute_value(el, 'table', 'number-columns-repeated') || 1).to_i
        curr_coli > coli
      end
      unless cellnode.nil? 
        return cellnode
      else
        add_cell
      end
    end
  end
  def add_cell(repeated=1)
    cell = Cell.new(self,first_unused_column_index)
    Tools.set_ns_attribute(cell.xmlnode,'table','number-columns-repeated',repeated) if repeated>1
    @xmlnode << cell.xmlnode
    cell
  end
  def used_range
    fu = first_unused_column_index
    (fu>1) ? 1..fu : nil
  end
  def first_unused_column_index
    1 + @xmlnode.elements.select{|n| n.name=='table-cell'}.reduce(0) do |sum, el|
      sum + (Tools.get_ns_attribute_value(el, 'table', 'number-columns-repeated') || 1).to_i
    end
  end
end

class RowGroup < RowWithXMLNode
  @readonly = :yes_always
  attr_reader :range
  attr_accessor :parent_array, :xmlnode
  def initialize(aparent_array,arange,axmlnode=nil)
    @parent_array = aparent_array
    @range = arange
    if axmlnode.nil?
      axmlnode = Row.empty_row_node
      Tools.set_ns_attribute(axmlnode,'table','number-rows-repeated',range.size) if range.size>1
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
  def range=(arange)
    @range=arange
    Tools.set_ns_attribute(@xmlnode,'table','number-rows-repeated',range.size, 1)
  end
end

class SingleRow < RowWithXMLNode
  @readonly = :no
  attr_accessor :xmlnode
  # index  Integer
  def initialize(aparent_array,aindex,axmlnode=nil)
    @parent_array = aparent_array
    @index = aindex
    if axmlnode.nil?
      axmlnode = Row.empty_row_node
    end
    @xmlnode = axmlnode
  end
  def self.new_from_rowgroup(rg)
    anode = rg.xmlnode
    Tools.remove_ns_attribute(anode,'table','number-rows-repeated')
    
    SingleRow.new(rg.parent_array,rg.range.begin,anode)
  end
  def normalize; self end
  def repeated?; false end
  def repeated; 1 end
  def range; (@index..@index) end
  def detach; true end
  def row; @index end
  
end

class LazyDetachableRow < Row
  @readonly = :yes_but_detachable
  def initialize(rowi)
    @index = rowi.to_i
  end
  def add_cell; detach.add_cell end
  def style_name=(value); detach.style_name=value end
  def row; @index end
end

## there are not data in this object, they are taken from RowGroup, but this is only readonly
class MemberOfRowGroup < LazyDetachableRow
  @readonly = :yes_but_detachable
  extend Forwardable
  delegate [:repeated?, :repeated, :xmlnode, :parent_array] => :@row_group
  attr_accessor :row_group # for dubugging

  # @index  Integer
  # @row_group   RepeatedRow
  def initialize(arowi,arow_group)
    super(arowi)
    @row_group = arow_group
    raise 'Wrong parameter given' unless @row_group.is_a? RowGroup
  end
  def detach  # detaches MemberOfRowGroup from its RowGroup perhaps splitting RowGroup
    @row_group.parent_array.detach(@index)
  end
  def cells(coli)
    Cell.new(self,coli,@row_group.cellnodes(coli)).tap{|n| n.mode = :repeated}
  end
  def nonemptycells
    @row_group.nonemptycellsindexes.collect{ |coli| cells(coli) }
  end
end

## this is a row outside the used bounds. the main purpose of this object is to magically synchronize to existing data, once they are created
class UninitializedEmptyRow < LazyDetachableRow
  @readonly = :yes_but_detachable
  attr_reader :parent_array  # debug only
  def initialize(aparent_array,arowi)
    super(arowi)
    @parent_array = aparent_array
  end
  def cells(coli)
    if still_out_of_used_range?
      Cell.new(self,coli,Cell.empty_cell_node).tap{|n| n.mode = :outbound}
    else
      @parent_array.get_row(@index).cells(coli)
    end
  end
  def normalize
    if still_out_of_used_range?
      self
    else
      @parent_array.get_row(@index)
    end
  end
  def detach; @parent_array.detach_of_bound_row_group(@index) end
  def still_out_of_used_range?; @index >= @parent_array.first_unused_row_index end
  def xmlnode; Row.empty_row_node end
  def nonemptycells; [] end
end

end