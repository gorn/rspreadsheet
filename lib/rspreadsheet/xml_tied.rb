module Rspreadsheet

# @private
class XMLTied
  def xml
    xmlnode.to_s
  end
end
  
# @private
# abstrac class. All successort MUST implement: set_index,xml_options,parent,index
class XMLTiedItem < XMLTied
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
    parentnode = parent.xmlnode
    if parentnode.nil?
      nil
    else
      parent.find_my_subnode_respect_repeated(index, xml_options)
    end
  end
  def detach_if_needed
    detach if repeated? # item did not exist individually yet, detach it within its parent and therefore make it individally editable
  end
  def detach
    parent.detach_if_needed if parent.respond_to?(:detach_if_needed)
    parent.detach_my_subnode_respect_repeated(index, xml_options)
    self
  end
  def _shift_by(diff)
    set_index(index + diff)
  end
  def range
    parent.find_my_subnode_range_respect_repeated(index,xml_options)
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
#     self.instance_variables.each do |variable|
#       instance_variable_set(variable,nil)
#     end
  end
  def delete
    parent.delete_subitem(index)
    invalidate_myself
  end

end

# abstract class. All importers MUST implement: prepare_subitem (and delete)
# terminology
#   item, subitem is object from @itemcache (quite often subclass of XMLTiedItem)
#   node, subnode is LibXML::XML::Node object
#
# this class is made to be included, not subclassed - the reason is in delete method which calls super
# @private

module XMLTiedArray
  attr_reader :itemcache

  def find_my_subnode_range_respect_repeated(aindex, options)
    index = 0
    xmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      if index+repeated >= aindex
        return (index+1..index+repeated)
      else
        index = index+repeated
      end
    end
    return (index+1..Float::INFINITY)
  end
  
  # vrátí xmlnode na souřadnici aindex
  def find_my_subnode_respect_repeated(aindex, options)
    find_subnode_respect_repeated(xmlnode,aindex, options)
  end
  # vrátí item na souřadnici aindex
  def subitem(aindex)
    aindex = aindex.to_i
    if aindex.to_i<=0
      raise 'Item index should be greater then 0' if Rspreadsheet.raise_on_negative_coordinates
      nil 
    else 
      @itemcache[aindex] ||= prepare_subitem(aindex)
    end
  end
  
  def subitems(*params)
    case params.length 
      when 0 then subitems_array
      when 1 then subitem(params[0]) 
      else raise Exception.new('Wrong number of arguments.')
    end
  end
  
  def subitems_array
    (1..self.size).collect do |i|
      subitem(i)
    end
  end
  
  def size; find_first_unused_index_respect_repeated(subitem_xml_options)-1 end
      
  def find_subnode_respect_repeated(axmlnode, aindex, options)
    result1, result2 = find_subnode_with_range_respect_repeated(axmlnode, aindex, options)
    return result1
  end  
    
  def find_subnode_with_range_respect_repeated(axmlnode, aindex, options)
    index = 0
    axmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      oldindex = index
      index = index+repeated
      if index>= aindex
        return node, oldindex..index
      end
    end
    return nil, index..Float::INFINITY
  end
  
  def prepare_repeated_subnode(times_repeated,options)
    result = LibXML::XML::Node.new(options[:xml_items_node_name],nil, Tools.get_namespace('table'))
    Tools.set_ns_attribute(result,'table',options[:xml_repeated_attribute],times_repeated, 1)
    result
  end
  
  def clone_before_and_set_repeated_attribute(node,times_repeated,options)
    newnode = node.copy(true)
    Tools.set_ns_attribute(newnode,'table',options[:xml_repeated_attribute],times_repeated,1)
    node.prev = newnode
  end
  
  # detaches subnode with aindex  
  def detach_my_subnode_respect_repeated(aindex, options)
    axmlnode = xmlnode
    node,index_range = find_subnode_with_range_respect_repeated(axmlnode, aindex, options)
    if index_range.size > 1 # pokud potřebuje vůbec detachovat
      if !node.nil? # detach subnode
        [index_range.begin+1..aindex-1,aindex..aindex,aindex+1..index_range.end].reject {|range| range.size<1}.each do |range| # create new structure by cloning
          clone_before_and_set_repeated_attribute(node,range.size,options)
        end
        node.remove! # remove the original node
      else # add outbound xmlnode
        [index_range.begin+1..aindex-1,aindex..aindex].reject {|range| range.size<1}.each do |range|
          axmlnode << prepare_repeated_subnode(range.size, options)
        end
      end
    end
    return find_subnode_respect_repeated(axmlnode, aindex, options)
  end
  
  def insert_my_subnode_before_respect_repeated(aindex, options)
    axmlnode = xmlnode
    
    node,index_range = find_subnode_with_range_respect_repeated(axmlnode, aindex, options)
    
    if !node.nil? # found the node, now do the insert
      [index_range.begin+1..aindex-1,aindex..index_range.end].reject {|range| range.size<1}.each do |range| # split  original node by cloning
        clone_before_and_set_repeated_attribute(node,range.size,options)
      end
      clone_before_and_set_repeated_attribute(node.prev,1,options)         # insert new node
      node.remove!                                                         # remove the original node
    else # insert outbound xmlnode
      [index+1..aindex-1,aindex..aindex].reject {|range| range.size<1}.each do |range|
	axmlnode << XMLTiedArray.prepare_repeated_subnode(range.size, options)
      end  
    end
    return find_subnode_respect_repeated(axmlnode, aindex, options)
  end

  def delete_my_subnode_respect_repeated(aindex,options)
    detach_my_subnode_respect_repeated(aindex,options) #TODO: tohle neni uplne spravne, protoze to zanecha skupinu rozdelenou na dve casti
    subitem(aindex).xmlnode.remove!
  end
  
  def find_first_unused_index_respect_repeated(options)
    index = 0
    return 1 if xmlnode.nil?
    xmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      index = index+repeated
    end
    return index+1
  end

  def add_empty_subitem_before(aindex)
    @itemcache.keys.sort.reverse.select{|i| i>=aindex }.each do |i| 
      @itemcache[i+1]=@itemcache.delete(i)
      @itemcache[i+1]._shift_by(1)
    end
    insert_my_subnode_before_respect_repeated(aindex,subitem_xml_options)  # nyní vlož node do xml
    @itemcache[aindex] =  subitem(aindex)
  end
  
  # clean up item from xml (handle possible detachments) and itemcache. leave the object invalidation on the object
  # this should not be called from nowhere but XMLTiedItem.delete
  def delete_subitem(aindex)
    options = subitem_xml_options
    delete_my_subnode_respect_repeated(aindex,options)  # vymaž node z xml
    @itemcache.delete(aindex)
    @itemcache.keys.sort.select{|i| i>=aindex+1 }.each do |i| 
      @itemcache[i-1]=@itemcache.delete(i)
      @itemcache[i-1]._shift_by(-1)
    end
  end
  
  def delete
    @itemcache.each do |key,item| 
      item.delete   # delete item - this destroys its subitems, xmlnode and invalidates it
      @itemcache.delete(key)  # delete the entry from the hash, normally this would mean this ceases to exist, if user does not have reference stored somewhere. Of he does, the object is invalidated anyways
    end
    super # this for example for Row objects calls XMLTiedItem.delete because Row is subclass of XMLTiedItem
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

  # truncate the item completely, deleting all its subitems
  def truncate
    subitems.reverse.each{ |subitem| subitem.delete }  # reverse je tu jen kvuli performanci, aby to mazal zezadu
  end  
end

end 
