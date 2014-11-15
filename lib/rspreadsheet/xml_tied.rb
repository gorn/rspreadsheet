module Rspreadsheet

class XMLTiedItem
  def mode
   case
     when xmlnode.nil? then :outbound
     when repeated>1  then :repeated
     else :regular
   end
  end
  def repeated; (Tools.get_ns_attribute_value(xmlnode, 'table', xml_repeated_attribute) || 1 ).to_i end
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
  def shift_by(diff)
    set_index(index + diff)
  end
  def range
    parent.find_my_subnode_range_respect_repeated(index,xml_options)
  end
end

module XMLTiedArray
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
      nil 
    else 
      @itemcache[aindex] ||= prepare_subitem(aindex)
    end
  end
  
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

  def find_first_unused_index_respect_repeated(options)
    index = 0
    xmlnode.elements.select{|node| node.name == options[:xml_items_node_name]}.each do |node|
      repeated = (node.attributes[options[:xml_repeated_attribute]] || 1).to_i
      index = index+repeated
    end
    return index+1
  end

  def insert_subitem_before(aindex)
    insert_subitem_before_with_options(aindex,subitem_xml_options)
  end
  def insert_subitem_before_with_options(aindex,options)
    @itemcache.keys.sort.reverse.select{|i| i>=aindex }.each do |i| 
      @itemcache[i+1]=@itemcache.delete(i)
      @itemcache[i+1].shift_by(1)
    end
    insert_my_subnode_before_respect_repeated(aindex,options)
    @itemcache[aindex] =  subitem(aindex)
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

end

end 
