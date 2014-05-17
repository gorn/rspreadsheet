require "rspreadsheet/version"
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'

module Rspreadsheet

  class << self
    def new
       Workbook.new
    end  
  end
end
