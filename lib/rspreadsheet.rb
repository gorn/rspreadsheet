require 'rspreadsheet/version'
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'
require 'class_extensions'

module Rspreadsheet

  # makes creating new workbooks as easy as `Rspreadsheet.new` or `Rspreadsheet.open('filename.ods')
  def self.new(filename=nil)
    Workbook.new(filename)
  end  
  def self.open(filename)
    Workbook.new(filename)
  end
end
