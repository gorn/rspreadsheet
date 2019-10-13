require 'rspreadsheet/row'
require 'rspreadsheet/column'
require 'rspreadsheet/image'
require 'rspreadsheet/tools'
require 'helpers/class_extensions'
# require 'forwardable'

module Rspreadsheet

class Worksheet
  include XMLTiedArray_WithRepeatableItems
  attr_accessor :xmlnode
  # @!group XMLTiedArray related methods
  def subnode_options; {
    :node_name => 'table-row',
    :ignore_groupings => ['table-header-rows'],
    :repeated_attribute => 'number-rows-repeated'
  } end

  def initialize(xmlnode_or_sheet_name,workbook) # workbook is here ONLY because of inserting images - to find unique name - it would be much better if it should bot be there
    initialize_xml_tied_array
    # set up the @xmlnode according to parameter
    case xmlnode_or_sheet_name
      when LibXML::XML::Node
        @xmlnode = xmlnode_or_sheet_name
      when String
        @xmlnode = Tools.prepare_ns_node('table','table')
        self.name = xmlnode_or_sheet_name
      else raise 'Provide name or xml node to create a Worksheet object'
    end
  end

  # name of the worksheet
  # @returns [String] the name of the worksheet
  def name
    Tools.get_ns_attribute_value(@xmlnode,'table','name')
  end

  def name=(value)
    Tools.set_ns_attribute(@xmlnode,'table','name', value)
  end

  def first_unused_row_index
    first_unused_subitem_index
  end

  def add_row_above(arowi)
    insert_new_empty_subitem_before(arowi)
  end

  def insert_cell_before(arowi,acoli)        # TODO: maybe move this to row level
    detach_row_in_xml(arowi)
    rows(arowi).insert_new_item(acoli)
  end

  def detach_row_in_xml(rowi)
    return detach_my_subnode_respect_repeated(rowi)
  end

  def nonemptycells
    used_rows_range.collect{ |rowi| rows(rowi).nonemptycells }.flatten
  end

  #@!group images
  def worksheet_images
    @worksheet_images ||= WorksheetImages.new(self)
  end

  def images_count
    worksheet_images.size
  end

  def images(*params)
    worksheet_images.subitems(*params)
  end

  def insert_image(filename,mime='image/png')
    worksheet_images.insert_image(filename,mime)
  end

  def insert_image_to(x,y,filename,mime='image/png')
    img = insert_image(filename,mime)
    img.move_to(x,y)
    img
  end

  #@!group XMLTiedArray_WithRepeatableItems connected methods
  def rows(*params)
    subitems(*params)
  end
  alias :row :rows # note that this is confusing terminology, deprecate?

  def prepare_subitem(rowi)
    Row.new(self,rowi)
  end

  def rowcache
    @itemcache
  end

  #@!group How to get to cells? (syntactic sugar)
  # Returns value of the cell given either by  row,column integer coordinates of by address.
  # @param [(Integer,Integer), String] either row and column of the cell (i.e. 3,5) or a string containing it address i.e. 'F12'
  def [](*params)
    cells(*params).andand.value
  end

  # Sets value of the cell given either by  row,column integer coordinates of by address.
  # It also sets the type of the cell according to type of the value. For details #see Cell.value=
  # This also allows syntax like
  #
  #      @sheet[1] = ['Jan', 'Feb', 'Mar']
  def []=(*params)
    if (params.size == 2) and params[0].kind_of?(Integer)
      rows(params[0]).cellvalues = params[1]
    else
      cells(*params[0..-2]).andand.value = params.last
    end
  end

  # Returns a Cell object placed in row and column or on a Cell on string address
  # @param [(Integer,Integer), String] either row and column of the cell (i.e. 3,5) or a string containing it address i.e. 'F12'
  def cells(*params)
    case params.length
      when 0 then raise 'Not implemented yet' #TODO: return list of all cells
      when 1..2
        r,c = Rspreadsheet::Tools.a2c(*params)
        arow = row(r)
        arow.andand.cell(c)
      else raise ArgumentError.new('Wrong number of arguments.')
    end
  end

  alias :cell :cells
  def column(param)
    _,coli = Rspreadsheet::Tools.a2c(1,param)
    Column.new(self,coli)
  end

  # Allows syntax like sheet.F15. TO catch errors easier, allows only up to three uppercase letters in colum part, althought it won't be necessarry to restrict.
  def method_missing method_name, *args, &block
    if method_name.to_s.match(/^([A-Z]{1,3})(\d{1,8})(=?)$/)
      row,col = Rspreadsheet::Tools.convert_cell_address_to_coordinates($~[1],$~[2])
      assignchar = $~[3]
      if assignchar == '='
        self.cells(row,col).value = args.first
      else
        self.cells(row,col).value
      end
    else
      super
    end
  end

  alias :rowcount :size
  def used_rows_range
    1..self.rowcount
  end
end

end
