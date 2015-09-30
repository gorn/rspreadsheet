require 'zip'
require 'libxml'

module Rspreadsheet
class Workbook
  attr_reader :filename
  attr_reader :xmlnode # debug
  def xmldoc; @xmlnode.doc end
  
  #@!group Worskheets methods
  def create_worksheet_from_node(source_node)
    sheet = Worksheet.new(source_node)
    register_worksheet(sheet)
    return sheet
  end
  def create_worksheet(name = "Strana #{worksheets_count}")
    sheet = Worksheet.new(name)
    register_worksheet(sheet)
    return sheet
  end
  # @return [Integer] number of sheets in the workbook
  def worksheets_count; @worksheets.length end
  # @return [String] names of sheets in the workbook
  def worksheet_names; @worksheets.collect{ |ws| ws.name } end
  # @param [Integer,String]
  # @return [Worskheet] worksheet with given index or name
  def worksheets(index_or_name)
    case index_or_name
      when Integer then begin
        case index_or_name
          when 0 then nil 
          when 1..Float::INFINITY then @worksheets[index_or_name-1]
          when -Float::INFINITY..-1 then @worksheets[index_or_name]    # zaporne indexy znamenaji pocitani zezadu
        end  
      end
      when String then @worksheets.select{|ws| ws.name == index_or_name}.first
      when NilClass then nil
      else raise 'method worksheets requires Integer index of the sheet or its String name'
    end
  end
  def [](index_or_name); self.worksheets(index_or_name) end
  #@!group Loading and saving related methods  
  def initialize(afilename=nil)
    @worksheets=[]
    @filename = afilename
    @content_xml = Zip::File.open(@filename || File.dirname(__FILE__)+'/empty_file_template.ods') do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @xmlnode = @content_xml.find_first('//office:spreadsheet')
    @xmlnode.find('./table:table').each do |node|
      create_worksheet_from_node(node)
    end
  end
  # @param [String] Optional new filename
  # Saves the worksheet. Optionally you can provide new filename.
  def save(new_filename=nil)
    if @filename.nil? and new_filename.nil? then raise 'New file should be named on first save.' end
    # if the filename has changed than first copy the original file to new location (or template if it is a new file)
    if new_filename
      FileUtils.cp(@filename || File.dirname(__FILE__)+'/empty_file_template.ods', new_filename)
      @filename = new_filename
    end
    Zip::File.open(@filename) do |zip|
      # it is easy, because @xmlnode in in sync with contents all the time
      zip.get_output_stream('content.xml') do |f|
        f.write @content_xml.to_s(:indent => false)
      end
    end
  end
  # @return Mime of the file
  def mime; 'application/vnd.oasis.opendocument.spreadsheet' end
  # @return [String] Prefered file extension
  def mime_preferred_extension; 'ods' end
  alias :mime_default_extension :mime_preferred_extension
  
  private 
  def register_worksheet(worksheet)
    index = worksheets_count+1
    @worksheets[index-1]=worksheet
    @xmlnode << worksheet.xmlnode if worksheet.xmlnode.doc != @xmlnode.doc
  end
end
end
