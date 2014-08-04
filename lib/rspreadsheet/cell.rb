require 'andand'

module Rspreadsheet

class Cell
  attr_reader :col, :xmlnode
  attr_reader  :parent_row  # for debug only
  attr_accessor :mode
  def self.empty_cell_node
    LibXML::XML::Node.new('table-cell',nil, Tools.get_namespace('table'))
  end
  def initialize(aparent_row,coli,asource_node=nil)
    raise "First parameter should be Row object not #{aparent_row.class}" unless aparent_row.kind_of?(Rspreadsheet::Row)
    @mode = :regular
    @parent_row = aparent_row
    if asource_node.nil?
      asource_node = Cell.empty_cell_node
    end
    @xmlnode = asource_node
    @col = coli
  end
  def ns_table; @parent_row.xmlnode.doc.root.namespaces.find_by_prefix('table') end
  def ns_office; @parent_row.xmlnode.doc.root.namespaces.find_by_prefix('office') end
  def ns_text; @parent_row.xmlnode.doc.root.namespaces.find_by_prefix('text') end
  def to_s; value end
  def xml; self.source_node.to_s; end
  def value_xml; self.source_node.children.first.children.first.to_s; end
  def coordinates; [row,col]; end
  def row; @parent_row.row; end
  def value
    gt = guess_cell_type
    if (@mode == :regular) or (@mode == @repeated)
      case 
        when gt == nil then nil
        when gt == Float then @xmlnode.attributes['value'].to_f
        when gt == String then @xmlnode.elements.first.andand.content.to_s
        when gt == Date then Date.strptime(@xmlnode.attributes['date-value'].to_s, '%Y-%m-%d')
        when gt == 'percentage' then @xmlnode.attributes['value'].to_f
      end
    elsif @mode == :outbound # needs to check whether it is still unbound
      if parent_row.still_out_of_used_range?
        nil
      else
        parent_row.normalize.cells(@col).value
      end  
    else
      raise "Unknown cell mode #{@mode}"
    end
  end
  def value=(avalue)
    if @mode == :regular
      gt = guess_cell_type(avalue)
      case
        when gt == nil then raise 'This value type is not storable to cell'
        when gt == Float 
          set_type_attribute('float')
          remove_all_value_attributes_and_content(@xmlnode)
          Tools.set_ns_attribute(@xmlnode,'office','value', avalue.to_s) 
          @xmlnode << LibXML::XML::Node.new('p', avalue.to_f.to_s, ns_text)
        when gt == String then
          set_type_attribute('string')
          remove_all_value_attributes_and_content(@xmlnode)
          @xmlnode << LibXML::XML::Node.new('p', avalue.to_s, ns_text)
        when gt == Date then 
          set_type_attribute('date')
          remove_all_value_attributes_and_content(@xmlnode)
          Tools.set_ns_attribute(@xmlnode,'office','date-value', avalue.strftime('%Y-%m-%d'))
          @xmlnode << LibXML::XML::Node.new('p', avalue.strftime('%Y-%m-%d'), ns_text) 
        when gt == 'percentage'
          set_type_attribute('float')
          remove_all_value_attributes_and_content(@xmlnode)
          Tools.set_ns_attribute(@xmlnode,'office','value', avalue.to_f.to_s) 
          @xmlnode << LibXML::XML::Node.new('p', (avalue.to_f*100).round.to_s+'%', ns_text)
      end
    elsif (@mode == :repeated) or (@mode == :outbound ) # Cell did not exist individually yet, detach row and create editable cell
      row = @parent_row.detach
      row.cells(@col).value = avalue
    else
      raise "Unknown cell mode #{@mode}"
    end
  end
  def remove_all_value_attributes_and_content(node)
    if att = Tools.get_ns_attribute(@xmlnode, 'office','value') then att.remove! end
    if att = Tools.get_ns_attribute(@xmlnode, 'office','date-value') then att.remove! end
    @xmlnode.content=''
  end
  def set_type_attribute(typestring)
    Tools.set_ns_attribute(@xmlnode,'office','value-type',typestring)
  end
  
  # based on @xmlnode and optionally value which is about to be assigned, guesses which type the result should be
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
      # if not succesfull then try guessing by type
      type = @xmlnode.attributes['value-type'].to_s
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
      
      result =
      if !typeguess.nil? # if not certain by value, but have a typeguess
        if !avalue.nil?  # with value we may try converting
          if (typeguess(avalue) rescue false) # if convertible then it is typeguess
            typeguess
          elsif (String(avalue) rescue false) # otherwise try string
            String
          else # if not convertible to anyhing concious then nil
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
  def inspect
    "#<Cell:[#{row},#{col}]=#{value}(#{guess_cell_type.to_s})"
  end
end

end

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


