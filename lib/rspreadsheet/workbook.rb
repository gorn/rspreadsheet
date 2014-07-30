require 'zip'
require 'libxml'

module Rspreadsheet
class Workbook
  attr_reader :worksheets, :filename
  attr_reader :xmlnode # debug
  def initialize(afilename=nil)
    @worksheets={}
    @filename = afilename
    content_xml = Zip::File.open(@filename || './lib/rspreadsheet/empty_file_template.ods') do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @xmlnode = content_xml.find_first('//office:spreadsheet')
    @xmlnode.find('./table:table').each do |node|
      create_worksheet_from_node(node)
    end
  end
  def save(new_filename=nil)
    if @filename.nil? and new_filename.nil? then raise 'New file should be named on first save.' end
    # if the filename has changed than first copy the original file to new location (or template if it is a new file)
    if new_filename
      FileUtils.cp(@filename || './lib/rspreadsheet/empty_file_template.ods', new_filename)
      @filename = new_filename
    end
    Zip::File.open(@filename) do |zip|
      # it is easy, because @content_xml in in sync with contents all the time
      zip.get_output_stream('content.xml') do |f|
        f.write @content_xml
      end
    end
  end
  def create_worksheet_from_node(source_node)
    sheet = Worksheet.new(source_node)
    register_worksheet(sheet)
    return sheet
  end
  def create_worksheet_with_name(name)
    sheet = Worksheet.new(name)
    register_worksheet(sheet)
    return sheet
  end
  def create_worksheet
    index = worksheets_count
    create_worksheet_with_name("Strana #{index}")
  end
  def register_worksheet(worksheet)
    index = worksheets_count+1
    @worksheets[index]=worksheet
    @worksheets[worksheet.name]=worksheet unless worksheet.name.nil?
    @xmlnode << worksheet.xmlnode if worksheet.xmlnode.doc != @xmlnode.doc
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
  def xmldoc
    @xmlnode.doc
  end
end
end
