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
    last.initialize_from_file(filename)
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
  
  def initialize_from_file(filename)
    # ověřit, zda soubor na disku existuje TODO: tady by to chtělo zobecnit na IO
    file = File.new(filename)
    # generate unique image name
    Image.get_unused_filename(file.extname)
    # copy picture to Pictures/ folder (within zip) under this name
    # change xml
  end

  def self.get_unused_filename(extension)
#     path = 'Pictures/'
#     filename_base = '11111111'
#     
#     puts zf.dir.entries('dir1').inspect
# 
#     
#     
#     iterator = ''
#     while  File.exist?(upload_path + filename_base + iterator.to_s + extension) or (!Document.find_by_filename(filename_base + iterator.to_s + extension).nil?)
#       if (iterator == '' )
#         iterator = 0
#         filename_base += '_'
#       end
#       iterator = iterator + 1
#     end
#     return filename_base + iterator.to_s + extension
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