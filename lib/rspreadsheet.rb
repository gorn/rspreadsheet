require 'rspreadsheet/version'
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'
require 'helpers/configuration'
# refinements
require 'helpers/class_extensions'
using ClassExtensions

module Rspreadsheet
  extend Configuration
  define_setting :raise_on_negative_coordinates, true

  # makes creating new workbooks as easy as `Rspreadsheet.new` or `Rspreadsheet.open('filename.ods')
  def self.new(filename=nil)
    Workbook.new(filename)
  end  
  def self.open(filename)
    Workbook.new(filename)
  end
end
