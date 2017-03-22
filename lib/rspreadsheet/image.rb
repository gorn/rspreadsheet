module Rspreadsheet
  
# Represents images embeded in a Worksheet
class WorksheetImages
  include XMLTiedArray
  def initialize(parent_worksheet)
    initialize_xml_tied_array
    @worksheet = parent_worksheet
  end
  
  def insert_image(filename,mime='image/png')
    if xmlnode.nil? #TODO: this needs to be solved more generally maybe on XMLTiedArray level
      @worksheet.xmlnode
    end
    push_new
    last.initialize_from_file(filename,mime)
  end
  
  # @!group XMLTiedArray_WithRepeatableItems related methods    
  def subitem_xml_options; {:xml_items_node_name => 'frame', :xml_items_node_namespace => 'draw'} end
  def prepare_subitem(index); Image.new(self,index) end
  def xmlnode; @worksheet.xmlnode.find('./table:shapes').first end
  def prepare_empty_xmlnode
    Tools.insert_as_first_node_child(
      @worksheet.xmlnode, 
      Tools.prepare_ns_node('table', 'shapes')
    )
  end
  def prepare_empty_subnode
    main_node = super # prepares <draw:frame/> node but it is entirely empty
    [
      ['draw', 'z-index', '1'], 
      ['draw', 'name', 'test'],
      ['draw', 'style-name', 'gr1'],
      ['draw', 'text-style-name', 'P1'],
      ['svg', 'width', '11.63mm'],
      ['svg', 'height', '10.83mm']
    ].each do |line|
      Tools.set_ns_attribute(main_node,line[0],line[1],line[2])
    end
    
    sub_node = Tools.prepare_ns_node('draw', 'image')
    [
      ['xlink', 'type', 'simple'],
      ['xlink', 'show', 'embed'],
      ['xlink', 'actuate', 'onLoad']
    ].each do |line|
      Tools.set_ns_attribute(sub_node,line[0],line[1],line[2])
    end
    
    sub_node << Tools.prepare_ns_node('text','p')
    main_node << sub_node
    main_node
  end

end

# Represents an image included in the spreadsheet. The Image can NOT exist
# "detached" from an spreadsheet
class Image < XMLTiedItem
  attr_reader :mime
  
  def initialize(worksheet,index)
    super(worksheet,index)
    @original_filename = nil
  end
  
  def initialize_from_file(filename,mime)
    # ověřit, zda soubor na disku existuje TODO: tady by to chtělo zobecnit na IO
    raise 'File does not exist or it is not accessible' unless File.exists?(filename)
    @original_filename = filename
    @mime = mime
    self
  end
  def xml_image_subnode
    xmlnode.find('./draw:image').first
  end
  
  def move_to(ax,ay)
    self.x = ax
    self.y = ay
  end
  
  def original_filename; @original_filename end
  
  def copy_to(ax,ay,worksheet)
    img = worksheet.insert_image_to(ax,ay,@original_filename)
    img.height = height
    img.width  = width
  end
  
  # TODO: put some sanity check for values into these
  def x=(value);      Tools.set_ns_attribute(xmlnode,'svg','x',     value) end
  def y=(value);      Tools.set_ns_attribute(xmlnode,'svg','y',     value) end
  def width=(value);  Tools.set_ns_attribute(xmlnode,'svg','width', value) end
  def height=(value); Tools.set_ns_attribute(xmlnode,'svg','height',value) end
  def name=(value);   Tools.set_ns_attribute(xmlnode,'draw','name', value) end
  def x;      Tools.get_ns_attribute_value(xmlnode,'svg','x') end
  def y;      Tools.get_ns_attribute_value(xmlnode,'svg','y') end
  def width;  Tools.get_ns_attribute_value(xmlnode,'svg','width') end
  def height; Tools.get_ns_attribute_value(xmlnode,'svg','height') end
  def name;   Tools.get_ns_attribute_value(xmlnode, 'draw', 'name', nil) end
  def internal_filename; Tools.get_ns_attribute_value(xml_image_subnode,'xlink','href')  end
  def internal_filename=(value)
    Tools.set_ns_attribute(xml_image_subnode,'xlink','href', value ) 
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