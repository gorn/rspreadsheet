require 'zip/zipfilesystem'
require 'libxml'

module Rspreadsheet
class Workbook
  attr_reader :worksheets, :filename
  def initialize afilename=nil
    @worksheets={}
    @filename = afilename
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
  def save(new_filename=nil)
    if @filename.nil? and new_filename.nil? then raise 'New file should be named on first save.' end
    # if the filename has changed than first copy the original file to new location (or template if it is a new file)
    if new_filename
      FileUtils.cp(@filename || './lib/rspreadsheet/empty_file_template.ods', new_filename)
      @filename = new_filename
    end
    Zip::ZipFile.open(@filename) do |zip|
      # it is easy, because @content_xml in in sync with contents all the time
      zip.get_output_stream('content.xml') do |f|
        f.write @content_xml
      end
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
