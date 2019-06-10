require 'rspreadsheet/version'
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'
require 'helpers/configuration'
# refinements
require 'helpers/class_extensions'

module Rspreadsheet
  extend Configuration

  define_setting :raise_on_negative_coordinates, true

  # makes creating new workbooks as easy as `Rspreadsheet.new` or `Rspreadsheet.open('filename.ods')
  def self.new(*params)
    raise ArgumentError.new("wrong number of arguments (given #{params.size}, expected 0-2)") if params.size >2 
  
    case params.last
      when Hash then options = params.pop 
      else options = {}
    end
    
    if options[:format].nil? # automatické heuristické rozpoznání formátu
      options[:format] = :standard
      unless params.first.nil?
        begin
          Zip::File.open(params.first)
        rescue
          options[:format] = :flat
        end
      end
    end
    
    case options[:format]
      when :flat , :fods then WorkbookFlat.new(*params)
      when :standard then Workbook.new(*params)
      else raise 'format of the file not recognized'
    end
  end  
  def self.open(filename, options = {})
    self.new(filename, options)
  end
end
