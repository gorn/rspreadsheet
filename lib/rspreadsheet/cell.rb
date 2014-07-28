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
  def value=(avalue); @value=avalue; self end
  def xml; self.source_node.to_s; end
  def value_xml; self.source_node.children.first.children.first.to_s; end
  def coordinates; [row,col]; end
  def row; @parent_row.row; end
  def value
    type = @xmlnode.attributes['value-type'].to_s
    if (@xmlnode.children.size == 0) and (not @xmlnode.attributes?)
      nil
    else
      case type
        when 'float' then @xmlnode.attributes['value'].to_f
        when 'string' then @xmlnode.elements.first.andand.content.to_s
        when 'date' then Date.strptime(@xmlnode.attributes['date-value'].to_s, '%Y-%m-%d')
        when 'percentage' then @xmlnode.attributes['value'].to_f
        else
          if @xmlnode.children.size == 0
            nil
          else 
            raise "Unknown type from #{@xmlnode.to_s} / children size=#{@xmlnode.children.size.to_s} / type=#{type}"
          end
        end
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


