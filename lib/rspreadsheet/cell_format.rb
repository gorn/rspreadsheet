# @markup markdown
# @author Jakub Tesinsky
# @title rspreadsheet Cell

module Rspreadsheet
using ClassExtensions if RUBY_VERSION > '2.1'

###
# Represents a format of a cell. This object is returned by `@cell.format` method and allows syntax like
#
#    @cell.format.bold = true
#    @cell.format.italic = false
#    @cell.format.color = '#45AC00'
#
# Also handles all logic for formats.
# @private
class CellFormat
  def initialize(cell)
    @cell = cell
  end
  def cellnode; @cell.xmlnode end
  
  # @!group text style attribute readers
  def bold;      get_text_style_node_attribute('font-weight') == 'bold' end
  alias :bold? :bold 
  def italic;    get_text_style_node_attribute('font-style') == 'italic' end
  def color;     get_text_style_node_attribute('color') end
  def font_size; get_text_style_node_attribute('font-size') end
  def background_color; get_cell_style_node_attribute('background-color') end
  def currency
    Tools.get_ns_attribute_value(cellnode,'office','currency',nil) 
  end
  
  # @!group text style attribute writers
  def bold=(value);     set_text_style_node_attribute('font-weight', value ? 'bold' : 'normal') end
  def italic=(value);   set_text_style_node_attribute('font-style',  value ? 'italic' : 'normal') end
  def color=(value);    set_text_style_node_attribute('color',  value) end
  def font_size=(value);set_text_style_node_attribute('font-size',  value) end
  def background_color=(value); set_cell_style_node_attribute('background-color',  value) end
    
  
  # @!group border related
  def top;    @top    ||= Border.new(self,:top)  end
  def bottom; @bottom ||= Border.new(self,:bottom) end
  def left;   @left   ||= Border.new(self,:left) end
  def right;  @right  ||= Border.new(self,:right) end
  alias :border_top :top
  alias :border_right :right
  alias :border_bottom :bottom
  alias :border_left :left
      
  # @!group manipulation of style related nodes
  def get_text_style_node_attribute(attribute_name)
    text_style_node.nil? ? nil : Tools.get_ns_attribute_value(text_style_node,'fo',attribute_name)
  end
  def get_cell_style_node_attribute(attribute_name)
    cell_style_node.nil? ? nil : Tools.get_ns_attribute_value(cell_style_node,'fo',attribute_name)
  end
  def set_text_style_node_attribute(attribute_name,value)
    make_sure_styles_are_not_shared
    if text_style_node.nil?
      self.create_text_style_node
      raise 'Text style node was not correctly initialized' if text_style_node.nil?
    end
    Tools.set_ns_attribute(text_style_node,'fo',attribute_name,value)
  end
  def set_cell_style_node_attribute(attribute_name,value)
    make_sure_styles_are_not_shared
    if cell_style_node.nil?
      self.create_cell_style_node
      raise 'Cell style node was not correctly initialized' if cell_style_node.nil?
    end
    Tools.set_ns_attribute(cell_style_node,'fo',attribute_name,value)
  end
  def make_sure_styles_are_not_shared  # with another cell (prevent issue #40)
    # if cell is reapeated, we must detach it first
    @cell.detach if @cell.mode != :regular   
    # if style node is shared, we must create our own first
    if style_shared?
      new_style_name = unused_cell_style_name
      
      new_style_node = style_node.copy(true)
      Tools.set_ns_attribute(new_style_node,'style','name',new_style_name)
      style_node.next = new_style_node
      self.style_name = new_style_name
      raise 'Unsucessfull unsharing of styles' if style_shared?
    end
  end

 # @!group initialization of style related nodes, if they do not exist
  def create_text_style_node
    create_style_node if style_name.nil? or style_node.nil?
    raise 'text_style_node already exists' unless text_style_node.nil?
    style_node << Tools.prepare_ns_node('style','text-properties')
  end
  def create_cell_style_node
    create_style_node if style_name.nil? or style_node.nil?
    raise 'cell_style_node already exists' unless cell_style_node.nil?
    style_node << Tools.prepare_ns_node('style','table-cell-properties')
  end
  def create_style_node
    if style_name.nil?
      proposed_style_name = unused_cell_style_name
      Tools.set_ns_attribute(cellnode,'table','style-name',proposed_style_name)
      raise 'Style name was not correctly initialized' if style_name!=proposed_style_name
    end
    anode =  Tools.prepare_ns_node('style','style')
    Tools.set_ns_attribute(anode, 'style', 'name', style_name)
    Tools.set_ns_attribute(anode, 'style', 'family', 'table-cell')
    Tools.set_ns_attribute(anode, 'style', 'parent-style-name', 'Default')
    automatic_styles_node << anode
    raise 'Style node was not correctly initialized' if style_node.nil?
  end
  
  # @!group other style node related routines
  def unused_cell_style_name
    last = (cellnode.nil? ? [] : cellnode.doc.root.find('./office:automatic-styles/style:style')).
      collect {|node| node['name']}.
      collect{ |name| /^ce(\d*)$/.match(name); $1.andand.to_i}.
      compact.max || 0
    "ce#{last+1}"
  end
  def automatic_styles_node; style_node_with_partial_xpath('') end
  def style_name; Tools.get_ns_attribute_value(cellnode,'table','style-name',nil) end
  def style_node; style_node_with_partial_xpath("/style:style[@style:name=\"#{style_name}\"]") end
  def text_style_node; style_node_with_partial_xpath("/style:style[@style:name=\"#{style_name}\"]/style:text-properties") end
  def cell_style_node; style_node_with_partial_xpath("/style:style[@style:name=\"#{style_name}\"]/style:table-cell-properties") end
  def style_node_with_partial_xpath(xpath)
    return nil if cellnode.nil?
    cellnode.doc.root.find("./office:automatic-styles#{xpath}").first 
  end
  
  # how many cells are using this style. repeated cells are counted only once
  def style_shared_count
    cellnode.doc.root.find("//table:table-cell[@table:style-name=\"#{style_name}\"]").size
  end
  def style_shared?
    style_shared_count>1
  end
  
  # @!group other routines
  def inspect
    "#<Rspreadsheet::CellFormat bold:#{bold?.inspect}, borders: #{top.get_border_string || 'none'}, #{right.get_border_string || 'none'}, #{bottom.get_border_string || 'none'}, #{left.get_border_string || 'none'}>"
  end
  
  # experimental. Allows to assign a style to cell. Currenty used only for automatic styles.
  def style_name=(value)
    Tools.set_ns_attribute(cellnode,'table','style-name',value)
  end

end

# represents one of the borders of a cell
class Border
  def initialize(cellformat,side)
    @cellformat = cellformat
    @side = side.to_s
    raise "Wrong side of border object, can be top, bottom, left or right" unless ['left','right','top','bottom'].include? @side
  end
  def cellnode; @cell.xmlnode end
  def attribute_name; "border-#{@side}" end
  
  def width=(value); set_border_string_part(1, value) end
  def style=(value); set_border_string_part(2, value.to_s) end
  def color=(value); set_border_string_part(3, value) end
  def width; get_border_string_part(1).to_f end
  def style; get_border_string_part(2) end
  def color; get_border_string_part(3) end
  def delete
    @cellformat.set_cell_style_node_attribute(attribute_name, 'none')
  end
  
  ## internals  
   
  # set part-th part of string which represents the border. String looks like "0.06pt solid #00ee00"
  # part is 1 for width, 2 for style or 3 for color
  def set_border_string_part(part,value)
    current_value = @cellformat.get_cell_style_node_attribute(attribute_name)
    
    if current_value.nil? or (current_value=='none')
      value_array = ['0.75pt', 'solid', '#000000']  # set default values
    else
      value_array = current_value.split(' ')  
    end
    raise 'Strange border attribute value. Does not have 3 parts' unless value_array.length == 3
    value_array[part-1]=value
    @cellformat.set_cell_style_node_attribute(attribute_name, value_array.join(' '))
  end
  
  def get_border_string
    result = @cellformat.get_cell_style_node_attribute(attribute_name) || @cellformat.get_cell_style_node_attribute('border')
    result = nil if result=='none'
    result
  end
  
  def get_border_string_part(part)
    border_string = get_border_string
    return nil if border_string.nil?
      
    value_array = border_string.split(' ')  
    raise 'Strange border attribute value. Does not have 3 parts' unless value_array.length == 3
    return value_array[part-1]
  end
  
end

end # module
















