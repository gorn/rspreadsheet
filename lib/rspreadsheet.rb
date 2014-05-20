require "rspreadsheet/version"
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'

module Rspreadsheet

  class << self
    def new
      Workbook.new
    end  
    def open(filename)
      Workbook.new
    end

  end
end
