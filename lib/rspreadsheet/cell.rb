# @markup markdown
# @author Jakub Tesinsky
# @title rspreadsheet Cell

require 'andand'
require 'rspreadsheet/xml_tied_item'
require 'date'
require 'time'
require 'bigdecimal'
require 'bigdecimal/util' # for to_d method
require 'helpers/class_extensions'

module Rspreadsheet
using ClassExtensions if RUBY_VERSION > '2.1'

###
# Represents a cell in spreadsheet which has coordinates, contains value, formula and can be formated.
# You can get this object like this (suppose that @worksheet contains {Rspreadsheet::Worksheet} object)
#
#     @worksheet.cells(5,2)
#
# Note that when using syntax like `@worksheet[5,2]` or `@worksheet.B5` you won't get this object, but rather the value of the cell.
# More precisely it is equvalient to @worksheet.cells(5,2).value.
# 

class Cell < XMLTiedItem
  attr_accessor :worksheet 
  attr_reader :rowi
  InternalDateFormat = '%Y-%m-%d'
  InternalTimeFormat = 'PT%HH%MM%SS'
#   InternalTimeFormat = 'PT%HH%MM%SS,%LS'
  
  # @!group XMLTiedItem related methods and extensions  
  def xml_options; {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'} end
  def parent; row end
  def coli; index end
    
  def set_rowi(arowi); @rowi = arowi end # this should ONLY be used by parent row
  def initialize(aworksheet,arowi,acoli)
    raise "First parameter should be Worksheet object not #{aworksheet.class}" unless aworksheet.kind_of?(Rspreadsheet::Worksheet)
    @worksheet = aworksheet
    @rowi = arowi
    initialize_xml_tied_item(row,acoli)
  end
  def row; @worksheet.rows(rowi) end
  def coordinates; [rowi,coli] end
  def to_s; value.to_s end
  def valuexml; self.valuexmlnode.andand.inner_xml end
  def valuexmlnode; self.xmlnode.elements.first end
  # use this to find node in cell xml. ex. xmlfind('.//text:a') finds all link nodes
  def valuexmlfindall(path)
    valuexmlnode.nil? ? [] : valuexmlnode.find(path)
  end
  def valuexmlfindfirst(path)
    valuexmlfindall(path).first
  end
  def inspect
    "#<Rspreadsheet::Cell\n row:#{rowi}, col:#{coli} address:#{address}\n type: #{guess_cell_type.to_s}, value:#{value}\n mode: #{mode}, format: #{format.inspect}\n>"
  end
  def value
    gt = guess_cell_type
    if (self.mode == :regular) or (self.mode == :repeated)
      case 
        when gt == nil then nil
        when gt == Float then xmlnode.attributes['value'].to_f
        when gt == String then xmlnode.elements.first.andand.content.to_s
        when gt == Date then Date.strptime(xmlnode.attributes['date-value'].to_s, InternalDateFormat)
        when gt == Time then self.time_value
        when gt == :percentage then xmlnode.attributes['value'].to_f
        when gt == :currency then xmlnode.attributes['value'].to_d
      end
    elsif self.mode == :outbound
      nil
    else
      raise "Unknown cell mode #{self.mode}"
    end
  end
  ## according to http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html#__RefHeading__1417674_253892949
  ## the value od time-value is in a "duration" format defined here https://www.w3.org/TR/xmlschema-2/#duration
  ## this method converts the time-value to Time object. Note that it does not check if the cell is in time-value
  ## or not, this is the responibility of caller
  def time_value
    Cell.parse_time_value(xmlnode.attributes['time-value'].to_s)
  end
  def self.parse_time_value(svalue)
    begin
      Time.strptime(svalue, 'PT%HH%MM%SS') # first iteration which does fail - see issue #23
    rescue
      begin
        Time.parse(svalue) # maybe add defaults for year-mont-day
      rescue  
        # ultimate manual parse
        m = /^P((?<years>[0-9]+)Y)?((?<months>[0-9]+)M)?((?<days>[0-9]+)D)?T((?<hours>[0-9]+)H)?((?<minutes>[0-9]+)M)?((?<seconds>[0-9]+(\.[0-9]+)?)S)$/ .match(svalue.delete(' '))
        
        seconds = m[:seconds].to_f
        minutes = m[:minutes].to_i
        hours = m[:hours].to_i
        days = m[:days].to_i
        months = m[:months].to_i
        years = m[:years].to_i
        
        if seconds >= 61 
          divisor = (seconds / 60).floor
          minutes,  = minutes + divisor
          seconds  = seconds - divisor * 60
        end
        
        if minutes >= 60
          divisor = (seconds / 60).floor
          minutes,  = minutes + divisor
          seconds  = seconds - divisor * 60
        end
        
        raise m[:years].inspect if m
      end  
    end
  end
  
  def value=(avalue)
    detach_if_needed
    if self.mode == :regular
      gt = guess_cell_type(avalue)
      case
        when gt == nil then raise 'This value type is not storable to cell'
        when gt == Float then
          remove_all_value_attributes_and_content(xmlnode)
          set_type_attribute('float')
          Tools.set_ns_attribute(xmlnode,'office','value', avalue.to_s) 
          xmlnode << Tools.prepare_ns_node('text','p', avalue.to_f.to_s)
        when gt == String then
          remove_all_value_attributes_and_content(xmlnode)
          set_type_attribute('string')
          xmlnode << Tools.prepare_ns_node('text','p', avalue.to_s)
        when gt == Date then 
          remove_all_value_attributes_and_content(xmlnode)
          set_type_attribute('date')
          avalue = avalue.strftime(InternalDateFormat)
          Tools.set_ns_attribute(xmlnode,'office','date-value', avalue)
          xmlnode << Tools.prepare_ns_node('text','p', avalue)
        when gt == Time then
          remove_all_value_attributes_and_content(xmlnode)
          set_type_attribute('time')
          Tools.set_ns_attribute(xmlnode,'office','time-value', avalue.strftime(InternalTimeFormat))
          xmlnode << Tools.prepare_ns_node('text','p', avalue.strftime('%H:%M'))
        when gt == :percentage then
          remove_all_value_attributes_and_content(xmlnode)
          set_type_attribute('percentage')
          Tools.set_ns_attribute(xmlnode,'office','value', '%0.2d%' % avalue.to_f) 
          xmlnode << Tools.prepare_ns_node('text','p', (avalue.to_f*100).round.to_s+'%')
      end
    else
      raise "Unknown cell mode #{self.mode}"
    end
  end
  def set_type_attribute(typestring)
    Tools.set_ns_attribute(xmlnode,'office','value-type',typestring)
    Tools.set_ns_attribute(xmlnode,'calcext','value-type',typestring)
  end
  def remove_all_value_attributes_and_content(node=xmlnode)
    if att = Tools.get_ns_attribute(node, 'office','value') then att.remove! end
    if att = Tools.get_ns_attribute(node, 'office','date-value') then att.remove! end
    if att = Tools.get_ns_attribute(node, 'office','time-value') then att.remove! end
    if att = Tools.get_ns_attribute(node, 'table','formula') then att.remove! end
    node.content=''
  end
  def remove_all_type_attributes
    set_type_attribute(nil)
  end
  def relative(rowdiff,coldiff)
    @worksheet.cells(self.rowi+rowdiff, self.coli+coldiff)
  end
  def type
    gct = guess_cell_type
    case 
      when gct == Float  then :float
      when gct == String then :string
      when gct == Date   then :date
      when gct == Time   then :time
      when gct == :percentage then :percentage
      when gct == :unassigned then :unassigned
      when gct == :currency then :currency
      when gct == NilClass then :empty
      when gct == nil then :unknown
      else :unknown
    end
  end
  def guess_cell_type(avalue=nil)
    # try guessing by value
    valueguess = case avalue
      when Numeric then Float
      when Time then Time
      when Date then Date
      when String,nil then nil
      else nil
    end
    result = valueguess

    if valueguess.nil? # valueguess is most important if not succesfull then try guessing by type from node xml
      typ = xmlnode.nil? ? 'N/A' : xmlnode.attributes['value-type']
      typeguess = case typ
        when nil then nil
        when 'float' then Float
        when 'string' then String
        when 'time' then Time
        when 'date' then Date
        when 'percentage' then :percentage
        when 'N/A' then :unassigned
        when 'currency' then :currency
        else 
          if xmlnode.elements.size == 0
            nil
          else 
            raise "Unknown type at #{coordinates.to_s} from #{xmlnode.to_s} / elements size=#{xmlnode.elements.size.to_s} / type=#{xmlnode.attributes['value-type'].to_s}"
          end
      end

      result =
      if !typeguess.nil? # if not certain by value, but have a typeguess
        if !avalue.nil?  # with value we may try converting
          if (typeguess(avalue) rescue false) # if convertible then it is typeguess
            typeguess
          elsif (String(avalue) rescue false) # otherwise try string
            String
          else # if not convertible to anything concious then nil
            nil 
          end
        else             # without value we just beleive typeguess
          typeguess
        end
      else # it not have a typeguess
        if (avalue.nil?) # if nil then nil
          NilClass
        elsif (String(avalue) rescue false) # convertible to String
          String
        else # giving up
          nil
        end
      end
    elsif valueguess == Float and xmlnode.andand.attributes['value-type'] == 'percentage'
      result = :percentage
    end
    result
  end
  def format
    @format ||= CellFormat.new(self)
  end
  def address
    Tools.convert_cell_coordinates_to_address(coordinates)
  end
  
  def formula
    rawformula = Tools.get_ns_attribute(xmlnode,'table','formula',nil).andand.value
    if rawformula.nil?
      nil 
    elsif rawformula.match(/^of:(.*)$/)
      $1
    else
      raise "Mischmatched value in table:formula attribute - does not start with of: (#{rawformula.to_s})"
    end
  end
  def formula=(formulastring)
    detach_if_needed
    raise 'Formula string must begin with "=" character' unless formulastring[0,1] == '='
    remove_all_value_attributes_and_content(xmlnode)
    remove_all_type_attributes
    Tools.set_ns_attribute(xmlnode,'table','formula','of:'+formulastring.to_s)
  end
  def blank?; self.type==:empty or self.type==:unassigned end
  
  def border_top;    format.border_top end
  def border_right;  format.border_right end
  def border_bottom; format.border_bottom end
  def border_left;   format.border_left end

end

# proxy object to allow cell.format syntax. Also handles all logic for formats.
# @private
class CellFormat
  def initialize(cell)
    @cell = cell
  end
  def cellnode; @cell.xmlnode end
  
  # text style attribute readers
  def bold;      get_text_style_node_attribute('font-weight') == 'bold' end
  alias :bold? :bold 
  def italic;    get_text_style_node_attribute('font-style') == 'italic' end
  def color;     get_text_style_node_attribute('color') end
  def font_size; get_text_style_node_attribute('font-size') end
  def get_text_style_node_attribute(attribute_name)
    text_style_node.nil? ? nil : Tools.get_ns_attribute_value(text_style_node,'fo',attribute_name)
  end
  def background_color; get_cell_style_node_attribute('background-color') end
  def get_cell_style_node_attribute(attribute_name)
    cell_style_node.nil? ? nil : Tools.get_ns_attribute_value(cell_style_node,'fo',attribute_name)
  end
  
  # text style attribute writers
  def bold=(value);     set_text_style_node_attribute('font-weight', value ? 'bold' : 'normal') end
  def italic=(value);   set_text_style_node_attribute('font-style',  value ? 'italic' : 'normal') end
  def color=(value);    set_text_style_node_attribute('color',  value) end
  def font_size=(value);set_text_style_node_attribute('font-size',  value) end
  def set_text_style_node_attribute(attribute_name,value)
    @cell.detach if @cell.mode != :regular
    if text_style_node.nil?
      self.create_text_style_node
      raise 'Style node was not correctly initialized' if text_style_node.nil?
    end
    Tools.set_ns_attribute(text_style_node,'fo',attribute_name,value)
  end
  def background_color=(value); set_cell_style_node_attribute('background-color',  value) end
  def set_cell_style_node_attribute(attribute_name,value)
    @cell.detach if @cell.mode != :regular
    if cell_style_node.nil?
      self.create_cell_style_node
      raise 'Style node was not correctly initialized' if cell_style_node.nil?
    end
    Tools.set_ns_attribute(cell_style_node,'fo',attribute_name,value)
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
    Tools.set_ns_attribute(anode, 'style', 'name', proposed_style_name)
    Tools.set_ns_attribute(anode, 'style', 'family', 'table-cell')
    Tools.set_ns_attribute(anode, 'style', 'parent-style-name', 'Default')
    automatic_styles_node << anode
    raise 'Style node was not correctly initialized' if style_node.nil?
  end
  
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
  def currency
    Tools.get_ns_attribute_value(cellnode,'office','currency',nil) 
  end
  #returns object representing top border of the cell
  def top;    @top    ||= Border.new(self,:top)  end
  def bottom; @bottom ||= Border.new(self,:bottom) end
  def left;   @left   ||= Border.new(self,:left) end
  def right;  @right  ||= Border.new(self,:right) end
  alias :border_top :top
  alias :border_right :right
  alias :border_bottom :bottom
  alias :border_left :left
  
  def inspect
    "#<Rspreadsheet::CellFormat bold:#{bold?.inspect}, borders:#{top.get_value_string.inspect} #{right.get_value_string.inspect} #{bottom.get_value_string.inspect} #{left.get_value_string.inspect}>"
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
   
  # set parth-th part of string which represents the border. String looks like "0.06pt solid #00ee00"
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
  
  def get_border_string_part(part)
    current_value = @cellformat.get_cell_style_node_attribute(attribute_name) || @cellformat.get_cell_style_node_attribute('border')
    if current_value.nil? or (current_value=='none')
      return nil
    else
      value_array = current_value.split(' ')  
      raise 'Strange border attribute value. Does not have 3 parts' unless value_array.length == 3
      return value_array[part-1]
    end
  end
  
  def get_value_string
    @cellformat.get_cell_style_node_attribute(attribute_name)
  end
  
end

end # module

















