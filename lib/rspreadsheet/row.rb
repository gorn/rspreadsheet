require 'rspreadsheet/cell'
require 'rspreadsheet/xml_tied_repeatable'

module Rspreadsheet

# Represents a row in a spreadsheet which has coordinates, contains value, formula and can be formated.
# You can get this object like this (suppose that @worksheet contains {Rspreadsheet::Worksheet} object)
#
#     @row = @worksheet.row(5)
#
# Mostly you will this object to access values of cells in the row
#
#     @row[2]           # identical to @worksheet[5,2] or @row.cells(2).value
#
# or directly row `Cell` objects
#
#     @row.cell(2)     # => identical to @worksheet.rows(5).cells(2)
#
# You can use it to manipulate rows
#
#     @row.add_row_above   # adds empty row above
#     @row.delete          # deletes row 
#
# and shifts all other rows down/up appropriatelly.

class Row < XMLTiedItem
  include XMLTiedArray_WithRepeatableItems
  ## @return [Worksheet] worksheet which contains the row
  # @!attribute [r] worksheet
  def worksheet; parent end
  ## @return [Integer] row index of the row
  # @!attribute [r] rowi
  def rowi; index end
  
  def initialize(aworksheet,arowi)
    initialize_xml_tied_array
    initialize_xml_tied_item(aworksheet,arowi)
  end
  
 # @!group Syntactic sugar  
  def cells(*params)
    if params.length == 1
      subitems(Tools.convert_column_name_to_index(params[0]))
    else
      subitems(*params)
    end
  end
  alias :cell :cells
  
  ## @return [String or Float or Date] value of the cell
  # @param coli [Integer ot String] colum index of the cell of colum letter
  # returns value of the cell at column `coli`. 
  #  
  #     @row = @worksheet.rows(5)     # with cells containing names of months
  #     @row[1]                       # => "January" 
  #     @row.cells(2).value           # => "February"  
  #     @row[1].class                 # => String
  def [](coli); cells(coli).value end
  # @param avalue [Array] 
  # sets value of cell in column `coli`
  def []=(coli,avalue); cells(coli).value=avalue end
  # @param avalue [Array] array with values 
  # sets values of cells of row to values from `avalue`. *Attention*: it deletes the rest of row
  #  
  #     @row = @worksheet.rows(5)     
  #     @row.values = ["January", "Feb", nil, 4] # =>  | January | Feb |  | 4 |
  #     @row[2] = "foo")                         # =>  | January | foo |  | 4 |
  def cellvalues=(avalue)
    self.truncate
    avalue.each_with_index{ |val,i| self[i+1] = val }
  end
  ## @return [Array
  # return array of cell values 
  # 
  #     @worksheet[3,3] = "text"
  #     @worksheet[3,1] = 123
  #     @worksheet.rows(3).cellvalues            # => [123, nil, "text"]
  def cellvalues
    cells.collect{|c| c.value}
  end
 # @!endgroup
  
 # @!group Other methods
  def style_name=(value); 
    detach_if_needed
    Tools.set_ns_attribute(xmlnode,'table','style-name',value)
  end
  def nonemptycells
    nonemptycellsindexes.collect{ |index| subitem(index) }
  end
  def nonemptycellsindexes
    myxmlnode = xmlnode
    if myxmlnode.nil?
      []
    else
      worksheet.find_nonempty_subnode_indexes(myxmlnode, subnode_options)
    end
  end
  alias :used_range :range
  # Inserts row above itself (and shifts itself and all following rows down)
  def add_row_above
    parent.add_row_above(rowi)
  end
  
  def next_row; relative(+1) end
  alias :next :next_row
  
  def relative(rowi_offset)
    worksheet.row(self.rowi+rowi_offset)
  end
  
 # @!group Private methods, which should not be called directly
  # @private
  # shifts internal represetation of row by diff. This should not be called directly
  # by user, it is only used by XMLTiedArray_WithRepeatableItems as hook when shifting around rows
  def _shift_by(diff)
    super
    @itemcache.each_value{ |cell| cell.set_rowi(rowi) }
  end
  
  # @!group XMLTiedArray_WithRepeatableItems related methods
  def subnode_options; {
    :node_name => 'table-cell', 
    :alt_node_names => ['covered-table-cell'], 
    :ignore_groupings => ['table-header-rows'], 
    :repeated_attribute => 'number-columns-repeated'
  } end
  def prepare_subitem(coli); Cell.new(worksheet,rowi,coli) end

end

end
