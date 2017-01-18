require 'helpers/class_extensions'

module Rspreadsheet

using ClassExtensions if RUBY_VERSION > '2.1'

# @private
class XMLTied
  def xml
    xmlnode.to_s
  end
end
  
# Abstract class representing and array which is tied to a particular element of XML file.
# It uses cashing to make access to array more effective. Implements the following methods:
#
#   * subitems(index) - returns subitem object on index 
#   * subitems - returns array of all subitems. Please note that first item is always nil so 
#     the array can be accessed using 1-based indexes.
#   
# Importer must provide:
#
#   * prepare_subitem(aindex) - must return newly created object representing item on aindex
#   * delete - ???
#   * xmlnode - must return xmlnode to which the array is tied. If speed is not a concern, 
#               consider not cashing it into variable, but finding it through document or parent. 
#               This prevents "broken" links. Sometimes when array is empty, the node does note
#               necessarily exists. That is fine, XMLTiedArray behaves correctly even with nil xmlnode,
#               of course util you want to insert something. If this may happens, importer must
#               provide method prepare_empty_xmlnode which prepares (and returns) empty xml node.
#               It is lazy called, as late as possible.
#   * subitem_xml_options - returns hash of options used to locate subitems in xml (TODO: rewrite in clear way)
#   * intilize must call initialize_xml_tied_array
#
# Terminology
#   * item, subitem is object from @itemcache (quite often subclass of XMLTiedItem)
#   * node, subnode is LibXML::XML::Node object
#
# this class is made to be included.
#
# Note for developers:
# Beware that the implementation of methods needs to be done in a way that it continues to
# work when items are "repeatable" - see XMLTiedArray_WithRepeatableItems. When impractical or impossible
# please implement the corresponding method in XMLTiedArray_WithRepeatableItems or at least override it there
# and make it raise exception.
# 
# @private
module XMLTiedArray
  attr_reader :itemcache
  
  def initialize_xml_tied_array
    @itemcache = Hash.new
  end
  
  # @!group accessing items
  
  # Returns item with index aindex
  def subitem(aindex)
    aindex = aindex.to_i
    if aindex.to_i<=0
      raise 'Item index should be greater then 0' if Rspreadsheet.raise_on_negative_coordinates
      nil 
    else 
      @itemcache[aindex] ||= prepare_subitem(aindex)
    end
  end
  
  def last
    subitem(size)
  end
  
  # Returns an array of subitems (when called without parameter) or an item on paricular index (when called with parameter). 
  def subitems(*params)
    case params.length 
      when 0 then subitems_array
      when 1 then subitem(params[0]) 
      else raise Exception.new('Wrong number of arguments.')
    end
  end

  # Returns array of subitems (repeated friendly)
  def subitems_array
    (1..self.size).collect do |i|
      subitem(i)
    end
  end
  
  # Number of subitems
  def size; first_unused_subitem_index-1 end
  alias :lenght :size
      
  # Finds first unused subitem index
  def first_unused_subitem_index
    (1 + xmlsubnodes.sum { |node| how_many_times_node_is_repeated(node) }).to_i
  end
  
  # @!group inserting new items  
  # Inserts empty subitem at the index position. Item currently on this position and all items 
  # after are shifter by index one.
  def insert_new_item(aindex)
    @itemcache.keys.sort.reverse.select{|i| i>=aindex }.each do |i| 
      @itemcache[i+1]=@itemcache.delete(i)
      @itemcache[i+1]._shift_by(1)
    end
    insert_new_empty_subnode_before(aindex)  # nyní vlož node do xml
    @itemcache[aindex] =  subitem(aindex)
  end
  alias :insert_new_empty_subitem_before :insert_new_item
  
  def push_new
    insert_new_item(first_unused_subitem_index)
  end
  
  # @!group other subitems methods
  # This is used (i.e. in first_unused_subitem_index) so it is flexible and can be reused in XMLTiedArray_WithRepeatableItems
  # @private
  def how_many_times_node_is_repeated(node); 1 end
  
#   # @!supergroup XML STRUCTURE internal handling methods #######################################

#   # @!group accessing subnodes
  # returns xmlnode with index
  # DOES not respect repeated_attribute
  def my_subnode(aindex)
    raise 'Using method which does not respect repeated_attribute with options that are using it. You probably donot want to do that.' unless subitem_xml_options[:xml_repeated_attribute].nil?
    return xmlsubnodes[aindex-1]
  end

  # @!group inserting new subnodes TODO: refactor out repeatable connected code
  def insert_new_empty_subnode_before(aindex)
    node_after = my_subnode(aindex)
    
    if !node_after.nil?
      node_after.prev = prepare_empty_subnode
      return node_after.prev
    elsif aindex==size+1
      # check whether xmlnode is ready for insetion
      if xmlnode.nil?
        prepare_empty_xmlnode
        if xmlnode.nil?
          raise 'Attempted call prepare_empty_xmlnode, but it did not created xmlnode correctly (it is still nil).'
        end
      end
      # do the insertion
      xmlnode <<  prepare_empty_subnode
      return xmlnode.last
    else
      raise IndexError.new("Index #{aindex} out of bounds (1..#{self.size})")
    end
  end    
 
  def prepare_empty_subnode
    Tools.prepare_ns_node(
      subitem_xml_options[:xml_items_node_namespace] || 'table',
      subitem_xml_options[:xml_items_node_name]
    )
  end
  
  # @!supergroup internal procedures dealing solely with xml structure ==========

  # importer must provide this only if it may happen that xmlnode is empty AND we will want to insert subitems
  def prepare_empty_xmlnode
    raise 'xmlnode is empty and I do not know how to create empty xmlnode. Please provide prepare_empty_xmlnode method in your object.'
  end
  
  # @!group finding and accessing subnodes
  # array containing subnodes of xmlnode which represent subitems
  def xmlsubnodes
    return [] if xmlnode.nil?
    ele = xmlnode.elements
    so = subitem_xml_options[:xml_items_node_name]
    
    ele.select do |node| 
      node.andand.name == so
    end
  end
    
end

end 
