require 'helpers/class_extensions'
require 'rspreadsheet/xml_tied_array' # BASTL to je jen kvuli XMLTied class

module Rspreadsheet

using ClassExtensions if RUBY_VERSION > '2.1'
  
# Note for developers: In case you want to represent a node containing many identical subnodes
# (like row contains cells, worksheet contains rows or images part includes images) than
#   1. Include module XMLTiedArray into parent class
#   2. Subclass XMLTiedItem by object which will represent individual items
#   3. Implement methods which are mentioned at both of these and the rest should work itself.
#   4. prepare_subitem calls method new of the item and usually sends index and parent + some
#      more so subitem can implement parent and index methods.

# @private
# abstract class. All successors have some options. MUST implement: 
#   * xml_options
#
# If you override intializer make sure you call initialize_xml_tied_item(aparent,aindex).
# 
# Optionally you may implement method which makes index accessible under more meaningful name
# like column or so. Optionally you may implement method which makes parent accessible under 
# more meaningful name like worksheet or so.
#
# By default parent is stored at initialization and never changed. If you do not want to cache
# it like this (to prevent inconsistencies) just override parent method to dfind the parent 
# dynamically (see Cell). If you do so you may want to disable the default parent handling by
# calling initialize_xml_tied_item(nil,index) in initializer (or do not call it at all if you
# have overridden index as well.
#
# Note: If there is a object representing parent (presumably using XMLTiedArray) than initialize
# signature must be reflected in parent prepare_subitem method
# 
class XMLTiedItem < XMLTied
  
  def initialize(aparent,aindex)
    initialize_xml_tied_item(aparent,aindex)
  end
  def parent; @xml_tied_parent end
  def index; @xml_tied_item_index end
  def set_index(aindex); @xml_tied_item_index=aindex end
  def index=(aindex); @xml_tied_item_index=aindex end
  
  # `xml_options[:xml_items_node_name]` gives the name of the tag representing cell
  # `xml_options[:number-columns-repeated]` gives the name of the previous tag which sais how many times the item is repeated
  def xml_options; abstract end
    
  def initialize_xml_tied_item(aparent,aindex)
    @xml_tied_parent = aparent unless aparent.nil?
    @xml_tied_item_index = aindex unless aindex.nil?
  end
    
  def mode
   case
     when xmlnode.nil? then :outbound
     when repeated>1  then :repeated
     else :regular
   end
  end
  def repeated; (Tools.get_ns_attribute_value(xmlnode, 'table', xml_options[:xml_repeated_attribute]) || 1 ).to_i end
  def repeated?; mode==:repeated || mode==:outbound end
  alias :is_repeated? :repeated?
  def xmlnode
    if parent.xmlnode.nil?
      nil
    else
      parent.my_subnode(index)
    end
  end
  def detach_if_needed
    detach if repeated? # item did not exist individually yet, detach it within its parent and therefore make it individally editable
  end
  def detach
    parent.detach_if_needed if parent.respond_to?(:detach_if_needed)
    parent.detach_my_subnode_respect_repeated(index)
    self
  end
  def _shift_by(diff)
    set_index(index + diff)
  end
  def range
    parent.my_subnode_range(index,xml_options)
  end
  def invalid_reference?; false end
  # destroys the object so it can not be used, this is necessarry to prevent
  # accessing cells and rows which has been long time ago deleted and do not represent
  # any physical object anymore
  def invalidate_myself
    raise_destroyed_cell_error = Proc.new {|*params| raise "Calling method of already destroyed Cell."}
    (self.methods - Object.methods + [:nil?]).each do |method| # "undefine" all methods
      self.singleton_class.send(:define_method, method, raise_destroyed_cell_error)
    end
    self.singleton_class.send(:define_method, :inspect, -> { "#<%s:0x%x destroyed cell>" % [self.class,object_id] })  # define descriptive inspect
    self.singleton_class.send(:define_method, :invalid_reference?, -> { true }) # define invalid_reference? method
    # invalidate variables
    @xml_tied_parent=nil
    @xml_tied_item_index=nil
#     self.instance_variables.each do |variable|
#       instance_variable_set(variable,nil)
#     end
  end
  def delete
    parent.delete_subitem(index)
    invalidate_myself
  end

end


end 
