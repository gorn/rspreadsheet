require 'andand'
require 'rspreadsheet/xml_tied'

module Rspreadsheet
  
class Cell < XMLTiedItem
  attr_accessor :worksheet, :coli, :rowi
  def xml_repeated_attribute;  'number-columns-repeated' end
  def xml_items_node_name; 'table-cell' end
  def xml_options; {:xml_items_node_name => xml_items_node_name, :xml_repeated_attribute => xml_repeated_attribute} end
  def parent; row end
  def index; @coli end
  def set_index(value); @coli=value end
  def set_rowi(arowi); @rowi = arowi end # this should ONLY be used by parent row
  def initialize(aworksheet,arowi,acoli)
    raise "First parameter should be Worksheet object not #{aworksheet.class}" unless aworksheet.kind_of?(Rspreadsheet::Worksheet)
    @worksheet = aworksheet
    @rowi = arowi
    @coli = acoli
  end
  def row; @worksheet.rows(rowi) end
  def coordinates; [rowi,coli] end
  def to_s; value.to_s end
  def valuexml; self.valuexmlnode.andand.inner_xml end
  def valuexmlnode; self.xmlnode.children.first end
  # use this to find node in cell xml. ex. xmlfind('.//text:a') finds all link nodes
  def valuexmlfindall(path)
    valuexmlnode.nil? ? [] : valuexmlnode.find(path)
  end
  def valuexmlfindfirst(path)
    valuexmlfindall(path).first
  end
  def inspect
    "#<Rspreadsheet::Cell:Cell\n row:#{rowi}, col:#{coli} address:#{address}\n type: #{guess_cell_type.to_s}, value:#{value}\n mode: #{mode}\n>"
  end
  def value
    gt = guess_cell_type
    if (self.mode == :regular) or (self.mode == :repeated)
      case 
        when gt == nil then nil
        when gt == Float then xmlnode.attributes['value'].to_f
        when gt == String then xmlnode.elements.first.andand.content.to_s
        when gt == Date then Date.strptime(xmlnode.attributes['date-value'].to_s, '%Y-%m-%d')
        when gt == 'percentage' then xmlnode.attributes['value'].to_f
      end
    elsif self.mode == :outbound
      nil
    else
      raise "Unknown cell mode #{self.mode}"
    end
  end
  def value=(avalue)
    detach_if_needed
    if self.mode == :regular
      gt = guess_cell_type(avalue)
      case
        when gt == nil then raise 'This value type is not storable to cell'
        when gt == Float 
          set_type_attribute('float')
          remove_all_value_attributes_and_content(xmlnode)
          Tools.set_ns_attribute(xmlnode,'office','value', avalue.to_s) 
          xmlnode << Tools.create_ns_node('text','p', avalue.to_f.to_s)
        when gt == String then
          set_type_attribute('string')
          remove_all_value_attributes_and_content(xmlnode)
          xmlnode << Tools.create_ns_node('text','p', avalue.to_s)
        when gt == Date then 
          set_type_attribute('date')
          remove_all_value_attributes_and_content(xmlnode)
          Tools.set_ns_attribute(xmlnode,'office','date-value', avalue.strftime('%Y-%m-%d'))
          xmlnode << Tools.create_ns_node('text','p', avalue.strftime('%Y-%m-%d')) 
        when gt == 'percentage'
          set_type_attribute('float')
          remove_all_value_attributes_and_content(xmlnode)
          Tools.set_ns_attribute(xmlnode,'office','value', avalue.to_f.to_s) 
          xmlnode << Tools.create_ns_node('text','p', (avalue.to_f*100).round.to_s+'%')
      end
    else
      raise "Unknown cell mode #{self.mode}"
    end
  end
  def set_type_attribute(typestring)
    Tools.set_ns_attribute(xmlnode,'office','value-type',typestring)
  end
  def remove_all_value_attributes_and_content(node)
    if att = Tools.get_ns_attribute(node, 'office','value') then att.remove! end
    if att = Tools.get_ns_attribute(node, 'office','date-value') then att.remove! end
    node.content=''
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
      when gct == :percentage then :percentage
      when gct == :unassigned then :unassigned
      when gct == nil then :unknown
      else :unknown
    end
  end
  def guess_cell_type(avalue=nil)
    # try guessing by value
    valueguess = case avalue
      when Numeric then Float
      when Date then Date
      when String,nil then nil
      else nil
    end
    result = valueguess

    if valueguess.nil? # valueguess is most important
      # if not succesfull then try guessing by type from node xml
      typ = xmlnode.nil? ? 'N/A' : xmlnode.attributes['value-type']
      typeguess = case typ
        when nil then nil
        when 'float' then Float
        when 'string' then String
        when 'date' then Date
        when 'percentage' then :percentage
        when 'N/A' then :unassigned
        else 
          if xmlnode.children.size == 0
            nil
          else 
            raise "Unknown type at #{coordinates.to_s} from #{xmlnode.to_s} / children size=#{xmlnode.children.size.to_s} / type=#{xmlnode.attributes['value-type'].to_s}"
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
        if (String(avalue) rescue false) # convertible to String
          String
        else
          nil
        end
      end
    end
    result
  end
  def format
    @format ||= CellFormat.new(self)
  end
  def address
    Tools.convert_cell_coordinates_to_address(coordinates)
  end
  def invalidate_references
    @worksheet = nil
    @coli = nil
    @rowi = nil
  end

end

# proxy object to allow cell.format syntax. Also handles all logic for formats.
class CellFormat
  attr_reader :bold
  def initialize(cell)
    @bold = false
    @cell = cell
  end
  def cellnode; @cell.xmlnode end
  def bold=(value)
    Rspreadsheet::Tools.set_ns_attribute(cellnode,'table','style-name','ce99')
  end
  def bold; Tools.get_ns_attribute_value(text_style_node,'fo','font-weight') == 'bold' end
  def italic; Tools.get_ns_attribute_value(text_style_node,'fo','font-style') == 'italic' end
  def color; Tools.get_ns_attribute_value(text_style_node,'fo','color') end
  def font_size; Tools.get_ns_attribute_value(text_style_node,'fo','font-size') end
  def background_color; Tools.get_ns_attribute_value(cell_style_node,'fo','background-color') end
 
  def unused_cell_style_name
    last = cellnode.doc.root.find('./office:automatic-styles/style:style').
      collect {|node| node['name']}.
      collect{ |name| /^ce(\d*)$/.match(name); $1.andand.to_i}.
      compact.max
    "ce#{last+1}"
  end
  def style_name; Tools.get_ns_attribute_value(cellnode,'table','style-name') end
  def style_node; cellnode.doc.root.find("./office:automatic-styles/style:style[@style:name=\"#{style_name}\"]").first end
  def text_style_node; cellnode.doc.root.find("./office:automatic-styles/style:style[@style:name=\"#{style_name}\"]/style:text-properties").first end
  def cell_style_node; cellnode.doc.root.find("./office:automatic-styles/style:style[@style:name=\"#{style_name}\"]/style:table-cell-properties").first end
end

end

















