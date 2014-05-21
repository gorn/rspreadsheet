require "rspreadsheet/version"
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'
require 'class_extensions'

module Rspreadsheet

  class << self
    def new filename=nil
      Workbook.new filename
    end  
    def open(filename)
      Workbook.new
    end
    def open
      
    end
  end
end
