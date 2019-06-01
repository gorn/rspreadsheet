if RUBY_VERSION > '2.1'

  module ClassExtensions

    refine Array do
      def sum(identity = 0, &block)
        if block_given?
          map(&block).sum(identity)
        else
          inject(0){ |sum, element| sum.to_f + element.to_f } || identity
        end
      end
    end
    
    refine LibXML::XML::Node do
      def equals?(node2); raise 'nic' end
      def ===(node2)  
        self.simplification_of?(node2) and node2.simplification_of?(self)
      end
      # if node2 contains at least all that I do
      def simplification_of?(node2)
        first_diff(node2).nil?
      end
      # return first difference where self has something more than node2 does
      def first_diff(node2)
        where = self.path.split('/').last
        
        return "#{where}> Equivalent node does not exist: #{self.name} != NOTHING" if node2.nil?
        return "#{where}> Names are different: #{self.name} != #{node2.name}" if (self.name != node2.name)
        self.attributes.each do |attr|
          return "#{where}> Attribute #{attr} have diffent values: #{attr.value} != #{node2.attributes[attr.name]}" unless node2.attributes[attr.name] == attr.value
        end

        elems1 = self.children
        elems2 = node2.children
        return "#{where}> elements have different number of children #{elems1.length} !=  #{elems2.length}" if (elems1.length != elems2.length)
        
        elems1.each_index do |i|
          raise "Nil for i=#{i}" if elems1[i].nil?
          if (elems1[i].node_type_name == 'text') 
            if elems2[i].nil? || (elems1[i].to_s) != (elems2[i].to_s)
              return  "#{where}> #{i+1}th text children are different: #{elems1[i].to_s} != #{elems2[i].to_s}"
            end
          elsif (elems1[i].node_type_name == 'element') 
            unless elems1[i].simplification_of?(elems2[i])
              return "#{where}/[#{i+1}]#{elems1[i].first_diff(elems2[i])}"
            end
          end
        end
          
        return nil
      end
      def elements
        result = []
        each_element { |e| result << e }
        return result
      end
    end
    
  end # module ClassExtensions

else # Monkeypatching

  class Array
    def sum(identity = 0, &block)
      if block_given?
        map(&block).sum(identity)
      else
        inject(0){ |sum, element| sum.to_f + element.to_f } || identity
      end
    end
  end

  class LibXML::XML::Node
    def ===(node2)  
      self.simplification_of?(node2) and node2.simplification_of?(self)
    end
    # if node2 contains at least all that I do
    def simplification_of?(node2)
      first_diff(node2).nil?
    end
    # return first difference where self has something more than node2 does
    def first_diff(node2)
      where = self.path.split('/').last
      
      return "#{where}> Equivalent node does not exist: #{self.name} != NOTHING" if node2.nil?
      return "#{where}> Names are different: #{self.name} != #{node2.name}" if (self.name != node2.name)
      self.attributes.each do |attr|
        return "#{where}> Attribute #{attr} have diffent values: #{attr.value} != #{node2.attributes[attr.name]}" unless node2.attributes[attr.name] == attr.value
      end

      elems1 = self.children
      elems2 = node2.children
      return "#{where}> elements have different number of children #{elems1.length} !=  #{elems2.length}" if (elems1.length != elems2.length)
      
      elems1.each_index do |i|
        raise "Nil for i=#{i}" if elems1[i].nil?
        if (elems1[i].node_type_name == 'text') 
          if elems2[i].nil? || (elems1[i].to_s) != (elems2[i].to_s)
            return  "#{where}> #{i+1}th text children are different: #{elems1[i].to_s} != #{elems2[i].to_s}"
          end
        elsif (elems1[i].node_type_name == 'element') 
          unless elems1[i].simplification_of?(elems2[i])
            return "#{where}/[#{i+1}]#{elems1[i].first_diff(elems2[i])}"
          end
        end
      end
        
      return nil
    end
    def elements
      result = []
      each_element { |e| result << e }
      return result
    end
    # TODO: posoudit jestli by to nebylo rychlejsi takto
#     def elements
#       children.select { |e| e.element? }
#     end
  end

end