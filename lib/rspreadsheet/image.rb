module Rspreadsheet
  
# Represents an image included in the spreadsheet.

class Images
  include XMLTiedArray
  def initialize(parent_worksheet)
    initialize_xml_tied_array
    @worksheet = parent_worksheet
  end
    
  def insert_image(filename)
    push_new
    last.filename = filename
  end

  # @!group XMLTiedArray_WithRepeatableItems related methods    
  def subitem_xml_options; {:xml_items_node_name => 'frame'} end
  def prepare_subitem(index); Image.new(self,index) end
  def xmlnode; @worksheet.xmlnode.find('./table:shapes').first end
end

class Image < XMLTiedItem
  def name
    Tools.get_ns_attribute_value(xmlnode, 'draw', 'name', nil)
  end
    
  # @!group XMLTiedItem related methods
  def xml_options; {:xml_items_node_name => 'frame'} end

#    
# Note: when creating new empty image we might have included xlink:type attribute but specification
# says it has default value simple [1] so we omit it. The same goes for 
# [1](http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html#attribute-xlink_type)
# 
end

end