class LibXML::XML::Node
  def elements
    result = []
    each_element { |e| result << e }
    return result
  end
  # if node2 contains at least all that I do
  def simpifation_of?(node2)
    return false if (self.name != node2.name)
    self.attributes.each do |attr|
      return false unless node2.attributes[attr.name] == attr.value
    end
    
    elems1 = self.elements
    elems2 = node2.elements
    return false if (elems1.length != elems2.length) 
    elems1.length.times do |i|
      unless 
        case elems1[i].node_type_name
          when 'text' 
            (elems1[i].to_s == elems2[i].to_s) 
          when 'element' 
            elems1[i].simpifation_of?(elems2[i])
          else true 
        end
      then 
        return false 
      end
    end
       
    return true
  end
  def equals?(node2)
    simpifation_of?(node2) and node2.simpifation_of?(self)
  end
end