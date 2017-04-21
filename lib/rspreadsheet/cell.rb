# @markup markdown
# @author Jakub Tesinsky
# @title rspreadsheet Cell

require 'andand'
require 'rspreadsheet/xml_tied_item'
require 'date'
require 'time'            # extended functions for time like Time.strptime
require 'bigdecimal'
require 'bigdecimal/util' # for to_d method
require 'helpers/class_extensions'
require 'rspreadsheet/cell_format'     # CellFormat and Border classes

module Rspreadsheet
using ClassExtensions if RUBY_VERSION > '2.1'

StartOfEpoch = Time.new(1899,12,30,0,0,0,0)

###
# Represents a cell in spreadsheet which has coordinates, contains value, formula and can be formated.
# You can get this object like this (suppose that @worksheet contains {Rspreadsheet::Worksheet} object)
#
#     @worksheet.cells(5,2)
#
# Note that when using syntax like `@worksheet[5,2]` or `@worksheet.B5` you won't get this object, but rather the value of the cell.
# More precisely it is equvalient to @worksheet.cells(5,2).value. Brief overview can be faound at [README]
# 

class Cell < XMLTiedItem
# RSpreadsheet::Worksheet in which the cell is contained.
  attr_accessor :worksheet 
# Row index of a cell. If you want to access the row object, see #row.  
  attr_reader :rowi
  
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
        when gt == :datetime then datetime_value
        when gt == :time then time_value
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
  ## or not, this is the responibility of caller. However beware that specification does not specify how the time 
  ## should be interpreted. By observing LibreOffice behaviour, I have found these options
  ##   1. "Time only cell" has time is stored as PT16H22M35S (16:22:35) where the duration is duration from midnight.  
  ##      Because ruby does NOT have TimeOfDay type we need to represent that as DateTime. I have chosen 1899-12-30 00:00:00 as 
  ##      StartOfEpoch time, because it plays well with case 2.
  ##   2. "DateTime converted to Time only cell" has time stored as PT923451H33M00.000000168S (15:33:00 with date part 2005-05-05 
  ##      before conversion to time only). It is strange format which seems to have hours meaning number of hours after 1899-12-30 00:00:00
  ##
  ## Returns time-value of the cell. It does not check if cell has or should have this value, it is responibility of caller to do so.
  def time_value
    Cell.parse_time_value(xmlnode.attributes['time-value'].to_s)
  end
  def self.parse_time_value(svalue)
    if (m = /^PT((?<hours>[0-9]+)H)?((?<minutes>[0-9]+)M)?((?<seconds>[0-9]+(\.[0-9]+)?)S)$/.match(svalue.delete(' ')))
      # time was parsed manually
      (StartOfEpoch + m[:hours].to_i*60*60 + m[:minutes].to_i*60 + m[:seconds].to_f.round(5))
      #BASTL: Rounding is here because LibreOffice adds some fractions of seconds randomly
    else
      begin
        Time.strptime(svalue, InternalTimeFormat)
      rescue
        Time.parse(svalue) # maybe add defaults for year-mont-day
      end
    end
  end
  def datetime_value
    vs = xmlnode.attributes['date-value'].to_s
    begin
      DateTime.strptime(vs, InternalDateTimeFormat)
    rescue
      begin
        DateTime.strptime(vs, InternalDateFormat)
      rescue
        DateTime.parse(vs)
      end
    end
  end
  def value=(avalue)
    detach_if_needed
    if self.mode == :regular
      gt = guess_cell_type(avalue)
#       raise 'here'+gt.to_s if avalue == 666.66
      case
        when gt == nil then raise 'This value type is not storable to cell'
        when gt == Float then
          remove_all_value_attributes_and_content
          set_type_attribute('float')
          Tools.set_ns_attribute(xmlnode,'office','value', avalue.to_s) 
          xmlnode << Tools.prepare_ns_node('text','p', avalue.to_f.to_s)
        when gt == String then
          remove_all_value_attributes_and_content
          set_type_attribute('string')
          xmlnode << Tools.prepare_ns_node('text','p', avalue.to_s)
        when gt == :datetime then 
          remove_all_value_attributes_and_content
          set_type_attribute('date')
          if avalue.kind_of?(DateTime) or avalue.kind_of?(Date) or avalue.kind_of?(Time)
            avalue = avalue.strftime(InternalDateTimeFormat)
            Tools.set_ns_attribute(xmlnode,'office','date-value', avalue)
            xmlnode << Tools.prepare_ns_node('text','p', avalue)
          end
        when gt == :time then
          remove_all_value_attributes_and_content
          set_type_attribute('time')
          if avalue.kind_of?(DateTime) or avalue.kind_of?(Date) or avalue.kind_of?(Time)
            Tools.set_ns_attribute(xmlnode,'office','time-value', avalue.strftime(InternalTimeFormat))
            xmlnode << Tools.prepare_ns_node('text','p', avalue.strftime('%H:%M'))
          end
        when gt == :percentage then
          remove_all_value_attributes_and_content
          set_type_attribute('percentage')
          Tools.set_ns_attribute(xmlnode,'office','value', '%0.2d%' % avalue.to_f) 
          xmlnode << Tools.prepare_ns_node('text','p', (avalue.to_f*100).round.to_s+'%')
        when gt == :currency then
          remove_all_value_attributes_and_content
          set_type_attribute('currency')
          unless avalue.nil?
            Tools.set_ns_attribute(xmlnode,'office','value', '%f' % avalue.to_d)
            xmlnode << Tools.prepare_ns_node('text','p', avalue.to_d.to_s+' '+self.format.currency)
          end
      end
    else
      raise "Unknown cell mode #{self.mode}"
    end
  end
  def set_type_attribute(typestring)
    Tools.set_ns_attribute(xmlnode,'office','value-type',typestring)
    Tools.set_ns_attribute(xmlnode,'calcext','value-type',typestring)
  end
  ## TODO: using this is NOT in line with the general intent of forward compatibility
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
      when gct == :datetime  then :datetime
      when gct == :time  then :time
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
      when Time then :time
      when Date, DateTime then :datetime
      when String,nil then nil
      else nil
    end
    result = valueguess
    
    if valueguess.nil?  # valueguess is most important if not succesfull then try guessing by type from node xml
      typ = xmlnode.nil? ? 'N/A' : xmlnode.attributes['value-type']
      typeguess = case typ
        when nil then nil
        when 'float' then Float
        when 'string' then String
        when 'time' then :time
        when 'date' then :datetime
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
      else  # it not have a typeguess
        if (avalue.nil?) # if nil then nil
          NilClass
        elsif (String(avalue) rescue false) # convertible to String
          String
        else # giving up
          nil
        end
      end
    elsif valueguess == Float 
      case xmlnode.andand.attributes['value-type'] 
        when 'percentage' then result = :percentage
        when 'currency' then result = :currency
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

  private
  InternalDateFormat = '%Y-%m-%d'
  InternalDateTimeFormat = '%FT%T'
  InternalTimeFormat = 'PT%HH%MM%SS'
end ## Cell

end # module
















