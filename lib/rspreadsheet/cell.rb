require 'andand'

module Rspreadsheet

class Cell
  attr_reader :col,:row, :xmlnode
  def initialize(aparent_row,coli,asource_node)
    raise "First parameter should be Row object not #{aparent_row.class}" unless aparent_row.kind_of?(Rspreadsheet::Row)
    @parent_row = aparent_row
    @xmlnode = asource_node
  end
  def to_s; value end
  def xml; self.source_node.to_s; end
  def value_xml; self.source_node.children.first.children.first.to_s; end
  def coordinates; [row,col]; end
  def row; @parent_row.row; end
  def value
    case guess_cell_type(@xmlnode)
      when nil then nil
      when Float then @xmlnode.attributes['value'].to_f
      when String then @xmlnode.elements.first.andand.content.to_s
      when Date then Date.strptime(@xmlnode.attributes['date-value'].to_s, '%Y-%m-%d')
      when 'percentage' then @xmlnode.attributes['value'].to_f
    end
  end
  def value=(avalue)
    case guess_cell_type(@xmlnode,avalue)
      when nil then raise 'This value type is not storable to cell'
      when Float 
        set_type_attribute('float')
        @xmlnode.attributes['office:value']=value.to_s
        @xmlnode.content=''
        @xmlnode << XML::Parser.string("<text:p>#{value}</text:p>").parse.root
      when String then @xmlnode.elements.first.andand.content.to_s
      when Date then Date.strptime(@xmlnode.attributes['date-value'].to_s, '%Y-%m-%d')
      when 'percentage' then @xmlnode.attributes['value'].to_f
    end
  end
  def set_type_attribute(typestring)
    @xmlnode.attributes['office:value-type']=typestring
  end
  
  # given cell xml node and optionally calue which is about to be assigned, guesses which type the result should be
  def guess_cell_type(axmlnode,avalue=nil)
    # try guessing by value
    valueguess = case avalue
      when Numeric then Float
      when Date then Date
      when String,nil then nil
      else nil
    end
    
    unless valueguess.nil? # valueguess is most important
      valueguess
    else # if not succesfull then try guessing by type
      type = axmlnode.attributes['value-type'].to_s
      typeguess = case type
        when 'float' then Float
        when 'string' then String
        when 'date' then Date
        when 'percentage' then 'percentage'
        else
          if @xmlnode.children.size == 0
            nil
          else 
            raise "Unknown type from #{@xmlnode.to_s} / children size=#{@xmlnode.children.size.to_s} / type=#{type}"
          end
        end
      end
      # if not certain by value, but value present, then try converting to typeguess
      if !avalue.nil? and !typeguess.nil?
        if (typeguess(avalue) rescue false) # if convertible
          typeguess
        elsif (String(avalue) rescue false)
          String
        else
          nil # not convertible to anyhing concious
        end
      else
        typeguess
      end
    end
  end
end

#   def former_initialize(arow,acol,source_node=nil)
#     @col = acol
#     @row = arow
#     @xmlnode = source_node
#     unless @xmlnode.nil?
#     end
#   end


# ## initialize cells
#   @cells = Hash.new do |hash, coords|
#     # we create empty cell and place it to hash, we do not have to check whether there is a cell in XML already, because it would be in hash as well
#     hash[coords]=Cell.new(coords[0],coords[1])
#     # TODO: create XML empty node here or upon save?
#   end
#   rowi = 1
#   unless @xmlnode.nil?
#     @xmlnode.elements.select{ |node| node.name == 'table-row'}.each do |row_source_node|
#       coli = 1
#       row_source_node.elements.select{ |node| node.name == 'table-cell'}.each do |cell_source_node|
#         initialize_cell(rowi,coli,cell_source_node)
#         coli += 1
#       end
#       rowi += 1
#     end
#   end


