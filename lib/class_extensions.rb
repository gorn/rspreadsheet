class LibXML::XML::Node
  def elements
    result = []
    each_element { |e| result << e }
    result
  end
end