module Rspreadsheet
  
# Represents an image included in the spreadsheet.

class Images
  include XMLTiedArray_WithRepeatableItems
  def initialize(parent_worksheet)
    super()
    @worksheet = parent_worksheet
  end
  def xmlnode; @worksheet.xmlnode.find('./table:shapes').first end

  # @!group XMLTiedArray_WithRepeatableItems related methods    
  def subitem_xml_options; {:xml_items_node_name => 'frame'} end
  def prepare_subitem(index); Image.new end

end

class Image
  def name; end
end
  
end