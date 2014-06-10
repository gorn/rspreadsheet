require 'andand'

module Rspreadsheet
class Cell
  attr_reader :value,:col,:row
  def initialize(arow,acol,source_node=nil)
    @col = acol
    @row = arow
    @source_node = source_node
    unless @source_node.nil?
      @type = @source_node.attributes['value-type'].to_s
      if (@source_node.children.size == 0) and (not @source_node.attributes?)
        @value = nil
      else
        # here you also ned to read style not only value
        @value = case @type
          when 'float'
            @source_node.attributes['value'].to_f
          when 'string'
            @source_node.elements.first.andand.content.to_s
          when 'date'
            Date.strptime(@source_node.attributes['date-value'].to_s, '%Y-%m-%d')
          when 'percentage'  
            @source_node.attributes['value'].to_f
          else
            if @source_node.children.size == 0
              nil
            else
              nil
#               raise "Unknown type from #{@source_node.to_s} / children size=#{@source_node.children.size.to_s} / type=#{@type}"
            end
        end
      end
    end
  end
  def to_s
    value
  end
  def value=(avalue)
    @value=avalue
    self
  end
  def coordinates
    [row,col]
  end
end
end
