class LibXML::XML::Node
  def elements
    result = []
    each_element { |e| result << e }
    result
  end
  def equals?(node2)
    self.eql?(node2)
    self.each_element do |subnode|
      puts subnode.xpath
    end
  end
end