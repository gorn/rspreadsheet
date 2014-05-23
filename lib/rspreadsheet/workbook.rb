# require 'zip/zipfilesystem'
require 'zip/zip'
require 'libxml'

module Rspreadsheet
class Workbook
  attr_reader :worksheets
  def initialize filename=nil
    @worksheets={}
    if filename.nil?
    else
      @content_xml = Zip::ZipFile.open(filename) do |zip|
        LibXML::XML::Document.io zip.get_input_stream('content.xml')
      end

      ndx = 0
      @content_xml.find_first('//office:spreadsheet').each_element { |node|
        sheet = Worksheet.new(node)
        @worksheets[ndx]=sheet
        @worksheets[node.name]=sheet
        ndx+=1
      }
    end
  end
  def create_worksheet(node=nil)
    sheet = Worksheet.new(node)
    @worksheets[worksheets_count]=sheet
    @worksheets[node.name]=sheet unless node.nil?
    return sheet
  end
  def worksheets
    @worksheets
  end
  def worksheets_count
    @worksheets.keys.select{ |k| k.kind_of? Numeric }.size #TODO: ?? max
  end
  def worksheet_names
    @worksheets.keys.reject{ |k| k.kind_of? Numeric }
  end
end
end
