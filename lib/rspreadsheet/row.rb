require 'rspreadsheet/cell'
require 'rspreadsheet/xml_tied'

# Currently this is only syntax sugar for cells and contains no functionality

module Rspreadsheet

class Row < XMLTiedItem
  include XMLTiedArray
  attr_reader :worksheet, :rowi, :cellcache
  def xml_repeated_attribute; 'number-rows-repeated' end
  def xml_items_node_name; 'table-row' end
  def xml_options; {:xml_items_node_name => xml_items_node_name, :xml_repeated_attribute => xml_repeated_attribute} end
  def subitem_xml_options; {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'} end
    
  def initialize(aworksheet,arowi)
    @worksheet = aworksheet
    @rowi = arowi
    @itemcache = Hash.new
  end
  def xmlnode; parent.find_my_subnode_respect_repeated(index, xml_options)  end
    
  # XMLTiedItem things and extensions
  def parent; @worksheet end
  def index; @rowi end
  def index=(value); @rowi=value end
    
  # XMLTiedArray rozšíření 
  def prepare_subitem(coli); Cell.new(@worksheet,@rowi,coli) end
  def cells(coli); subitem(coli) end
  def cellcache; @itemcache end
    
  # další
  def style_name=(value); 
    detach_if_needed
    Tools.set_ns_attribute(xmlnode,'table','style-name',value)
  end
  def nonemptycells
    nonemptycellsindexes.collect{ |index| subitem(index) }
  end
  def nonemptycellsindexes
    myxmlnode = xmlnode
    if myxmlnode.nil?
      []
    else
      @worksheet.find_nonempty_subnode_indexes(myxmlnode, {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'})
    end
  end
  alias :used_range :range
  def shift_by(diff)
    @rowi = @rowi + diff
  end
end

# class Row
#   def initialize
#     @readonly = :unknown
#     @cells = {}
#   end
#   def worksheet; @parent_array.worksheet end
#   def parent_array; @parent_array end  # for debug only
#   def used_col_range; 1..first_unused_column_index-1  end
#   def used_range; used_col_range  end
#   def first_unused_column_index; raise 'this should be redefined in subclasses' end
# end

  
#  -------------------------- 
  
  
# # XmlTiedArrayItemGroup is internal representation of repeated items in XmlTiedArray.
# class XmlTiedArrayItemGroup
# #   extend Forwardable
# #   delegate [:normalize ] => :@row_group
# 
#   def normalize; @rowgroup.normalize  end
#   def range; @rowgroup.range end
#   def repeated?; self.range.size>1 end
#   def xmlnode; @rowgroup.xmlnode end
# 
#   def initialize(aparent_array,arange,axmlnode=nil)
#     @rowgroup = RowGroup.new(aparent_array,arange,axmlnode)
#   end
#   def self.new_from_xml
#   end
#   def to_rowgroup
#     @rowgroup
#   end
#   def range=(arange)
#     
#   end
# end

# array which synchronizes with xml structure and reflects. number-xxx-repeated attributes
# also caches returned objects for indexes.
# options must contain
#   :xml_items, :xml_repeated_attribute, :object_type

# class XmlTiedArray < Array
#   def initialize(axmlnode, options={}) # TODO get rid of XmlTiedArray
#     @xmlnode = axmlnode
#     @options = options
#     
#     missing_options = [:xml_repeated_attribute,:xml_items_node_name,:object_type]-@options.keys
#     raise "Some options missing (#{missing_options.inspect})" unless missing_options.empty?
#     
#     unless @xmlnode.nil?
#       @xmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |group_source_node|
#         self << parse_xml_to_group(group_source_node) # it is in @xmlnode so suffices to add object to @rowgroups
#       end
#     end
#     @itemcache=Hash.new()
#   end
#   def parse_xml_to_group(size_or_xmlnode) # parses xml to new RowGroup which can be added at the end
#     # reading params
#     if size_or_xmlnode.kind_of? LibXML::XML::Node
#       size = (size_or_xmlnode[@options[:xml_repeated_attribute]] || 1).to_i
#       node = size_or_xmlnode
#     elsif size_or_xmlnode.to_i>0
#       size = size_or_xmlnode.to_i
#       node = nil
#     else
#       return nil
#     end
#     index = first_unused_index
#     # construct result
#     Rspreadsheet::XmlTiedArrayItemGroup.new(self,index..index+size-1,node)
#   end
#   def add_item_group(size_or_xmlnode)
#     result = parse_xml_to_group(size_or_xmlnode)
#     self << result
#     @xmlnode << result.xmlnode
#     result
#   end
#   def first_unused_index
#     empty? ? 1 : last.range.end+1
#   end
#   # prolonges the RowArray to cantain rowi and returns it
#   def detach_of_bound_item(index)
#     fill_row_group_size = index-first_unused_index
#     if fill_row_group_size>0
#       add_item_group(fill_row_group_size) 
#     end
#     add_item_group(1)
#     get_item(index)   # aby se odpoved nacacheovala
#   end
#   def get_item_group(index)
#     find{ |item_group| item_group.range.cover?(index) }
#   end
#   def detach_item(index); get_item(index) end # TODO předělat do lazy podoby, kdy tohle nebude stejny
#   def get_item(index)
#     if index>= first_unused_index
#       nil
#     else
#       @itemcache[index] ||= Rspreadsheet::XmlTiedArrayItem.new(self,index)
#     end
#   end
#   # This detaches item index from the group and perhaps splits the RowGroup
#   # into two pieces. This makes the row individually editable.
#   def detach(index)
#     group_index = get_group_index(index)
#     item_group = self[group_index]
#     range = item_group.range
#     return self if range==(index..index)
# 
#     # prepare new components
#     replaceby = []
#     replaceby << RowGroup.new(self,range.begin..index-1)
#     replaceby << (result = SingleRow.new(self,index))
#     replaceby << RowGroup.new(self,index+1..range.end)
#     
#     # put original range somewhere in replaceby and shorten it
#     
#     if index>range.begin
#       replaceby[0] = item_group
#       item_group.range = range.begin..index-1
#     else
#       replaceby[2] = item_group
#       item_group.range = index+1..range.end
#     end
#     
#     # normalize and delete empty parts
#     replaceby = replaceby.map(&:normalize).compact
#     
#     # do the replacement in xml
#     marker = LibXML::XML::Node.new('temporarymarker')
#     item_group.xmlnode.next = marker
#     item_group.xmlnode.remove!
#     replaceby.each{ |rg| 
#       marker.prev = rg.xmlnode
#     } 
#     marker.remove!
#     
#     # do the replacement in array
#     self[group_index..group_index]=replaceby
#     result
#   end
#   private
#   def get_group_index(index)
#     self.find_index{ |rowgroup| rowgroup.range.cover?(index) }
#   end
# end

# class XmlTiedArrayItem
#   attr_reader :index
#   def initialize(aarray,aindex)
#     @array = aarray
#     @index = aindex
#     if self.virtual?
#       @object = nil
#     else
#       @object = @array.options[:object_type].new(group.xmlnode)
#     end
#   end
#   def group; @array.get_item_group(index) end
#   def repeated?; group.repeated? end
#   def virtual?; ! self.repeated? end
#   def array
#     raise 'Group empty' if @group.nil? 
#     @array
#   end
# end

# class RowArray < XmlTiedArray
#   attr_reader :row_array_cache
#   def initialize(aworksheet,aworksheet_node)
#     @worksheet = aworksheet
#     @row_array_cache = Hash.new()
#     super(aworksheet_node, :xml_items_node_name => 'table-row', :xml_repeated_attribute => xml_repeated_attribute, :object_type=>Row)
#   end
#   def get_row(rowi)
#     if @row_array_cache.has_key?(rowi)
#       return @row_array_cache[rowi]
#     end
#     item = self.get_item(rowi)
#     @row_array_cache[rowi] = if item.nil?
#       if rowi>0 then Rspreadsheet::UninitializedEmptyRow.new(self,rowi) else nil end
#     else
#       if item.repeated?
# 	Rspreadsheet::MemberOfRowGroup.new(item.index, item.group.to_rowgroup)
#       else
# 	Rspreadsheet::SingleRow.new_from_rowgroup(item.group.to_rowgroup)
#       end
#     end
#   end
#   # aliases
#   def first_unused_row_index; first_unused_index end
#   def worksheet; @worksheet end
#   def detach_of_bound_row_group(index)
#     super(index)
#     return get_row(index)
#   end
# end

# class Row
#   def initialize
#     @readonly = :unknown
#     @cells = {}
#   end
#   def self.empty_row_node
#     LibXML::XML::Node.new('table-row',nil, Tools.get_namespace('table'))
#   end
#   def worksheet; @parent_array.worksheet end
#   def parent_array; @parent_array end  # for debug only
#   def used_col_range; 1..first_unused_column_index-1  end
#   def used_range; used_col_range  end
#   def first_unused_column_index; raise 'this should be redefined in subclasses' end
#   def cells(coli)
#     coli = coli.to_i
#     return nil if coli.to_i<=0
#     @cells[coli] ||= get_cell(coli)
#   end
# end

# class RowWithXMLNode < Row
#   attr_accessor :xmlnode
#   def style_name=(value); Tools.set_ns_attribute(@xmlnode,'table','style-name',value)  end
#   def get_cell(coli)
#     Cell.new(self,coli,cellnodes(coli))
#   end
#   def nonemptycells
#     nonemptycellsindexes.collect{ |index| cells(index) }
#   end
#   def nonemptycellsindexes
#     used_col_range.to_a.select do |coli|
#       cellnode = cellnodes(coli)
#       !(cellnode.content.nil? or cellnode.content.empty? or cellnode.content =='') or
#       !cellnode.attributes.to_a.reject{ |attr| attr.name == 'number-columns-repeated'}.empty?
#     end
#   end
#   def cellnodes(coli)
#     cellnode = nil
#     while true 
#       curr_coli=1
#       cellnode = @xmlnode.elements.select{|n| n.name=='table-cell'}.find do |el|
#         curr_coli += (Tools.get_ns_attribute_value(el, 'table', 'number-columns-repeated') || 1).to_i
#         curr_coli > coli
#       end
#       unless cellnode.nil? 
#         return cellnode
#       else
#         add_cell
#       end
#     end
#   end
#   def add_cell(repeated=1)
#     cell = Cell.new(self,first_unused_column_index)
#     Tools.set_ns_attribute(cell.xmlnode,'table','number-columns-repeated',repeated) if repeated>1
#     @xmlnode << cell.xmlnode
#     cell
#   end
#   def first_unused_column_index
#     1 + @xmlnode.elements.select{|n| n.name=='table-cell'}.reduce(0) do |sum, el|
#       sum + (Tools.get_ns_attribute_value(el, 'table', 'number-columns-repeated') || 1).to_i
#     end
#   end
# end

# class RowGroup < RowWithXMLNode
#   @readonly = :yes_always
#   attr_reader :range
#   attr_accessor :parent_array, :xmlnode
#   def initialize(aparent_array,arange,axmlnode=nil)
#     super()
#     @parent_array = aparent_array
#     @range = arange
#     if axmlnode.nil?
#       axmlnode = Row.empty_row_node
#       Tools.set_ns_attribute(axmlnode,'table','number-rows-repeated',range.size) if range.size>1
#     end
#     @xmlnode = axmlnode
#   end
#   # returns SingleRow if size of range is 1 and nil if it is 0 or less
#   def normalize
#     case range.size
#       when 2..Float::INFINITY then self
#       when 1 then SingleRow.new_from_rowgroup(self)
#       else nil
#     end
#   end
#   def repeated;  range.size   end
#   def repeated?; range.size>1 end
#   def range=(arange)
#     @range=arange
#     Tools.set_ns_attribute(@xmlnode,'table','number-rows-repeated',range.size, 1)
#   end
# end

# class SingleRow < RowWithXMLNode
#   @readonly = :no
#   attr_accessor :xmlnode
#   # index  Integer
#   def initialize(aparent_array,aindex,axmlnode=nil)
#     super()
#     @parent_array = aparent_array
#     @index = aindex
#     if axmlnode.nil?
#       axmlnode = Row.empty_row_node
#     end
#     @xmlnode = axmlnode
#   end
#   def self.new_from_rowgroup(rg)
#     anode = rg.xmlnode
#     Tools.remove_ns_attribute(anode,'table','number-rows-repeated')
#     SingleRow.new(rg.parent_array,rg.range.begin,anode)
#   end
#   def normalize; self end
#   def repeated?; false end
#   def repeated; 1 end
#   def range; (@index..@index) end
#   def detach; self end
#   def row; @index end
#   def still_out_of_used_range?; false end
# end

# class LazyDetachableRow < Row
#   @readonly = :yes_but_detachable
#   def initialize(rowi)
#     super()
#     @index = rowi.to_i
#   end
#   def add_cell; detach.add_cell end
#   def style_name=(value); detach.style_name=value end
#   def row; @index end
# end

# ## there are not data in this object, they are taken from RowGroup, but this is only readonly
# class MemberOfRowGroup < LazyDetachableRow
#   @readonly = :yes_but_detachable
#   extend Forwardable
#   delegate [:repeated?, :repeated, :xmlnode, :parent_array] => :@row_group
#   attr_accessor :row_group # for dubugging
# 
#   # @index  Integer
#   # @row_group   RepeatedRow
#   def initialize(arowi,arow_group)
#     super(arowi)
#     @row_group = arow_group
#     raise 'Wrong parameter given - class is '+@row_group.class.to_a unless @row_group.is_a? RowGroup
#   end
#   def detach  # detaches MemberOfRowGroup from its RowGroup perhaps splitting RowGroup
#     @row_group.parent_array.detach(@index)
#   end
#   def get_cell(coli)
#     c = Cell.new(self,coli,@row_group.cellnodes(coli))
#     c.mode = :repeated
#     c
#   end
#   def first_unused_column_index
#     @row_group.first_unused_column_index
#   end
#   def nonemptycells
#     @row_group.nonemptycellsindexes.collect{ |coli| cells(coli) }
#   end
# end


end