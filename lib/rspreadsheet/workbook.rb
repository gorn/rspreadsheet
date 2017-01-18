require 'zip'
require 'libxml'

module Rspreadsheet
class Workbook
  attr_reader :filename
  attr_reader :xmlnode # debug
  def xmldoc; @xmlnode.doc end
  
  #@!group Worskheets methods
  def create_worksheet_from_node(source_node)
    sheet = Worksheet.new(source_node,self)
    register_worksheet(sheet)
    return sheet
  end
  def create_worksheet(name = "Sheet#{worksheets_count+1}")
    sheet = Worksheet.new(name,self)
    register_worksheet(sheet)
    return sheet
  end
  alias :add_worksheet :create_worksheet 
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
  alias :worksheet :worksheets
  alias :sheet :worksheets
  alias :sheets :worksheets
  def [](index_or_name); self.worksheets(index_or_name) end
  #@!group Loading and saving related methods
  
  # @return Mime of the file
  def mime; 'application/vnd.oasis.opendocument.spreadsheet'.freeze end
  # @return [String] Prefered file extension
  def mime_preferred_extension; 'ods'.freeze end
  alias :mime_default_extension :mime_preferred_extension
  
  def initialize(afilename=nil)
    @worksheets=[]
    @filename = afilename
    @content_xml = Zip::File.open(@filename || TEMPLATE_FILE) do |zip|
      LibXML::XML::Document.io zip.get_input_stream(CONTENT_FILE_NAME)
    end
    @xmlnode = @content_xml.find_first('//office:spreadsheet')
    @xmlnode.find('./table:table').each do |node|
      create_worksheet_from_node(node)
    end
  end
  
  # @param [String] Optional new filename
  # Saves the worksheet. Optionally you can provide new filename.

  def save(new_filename_or_io_object=nil)
    @par = new_filename_or_io_object
    if @filename.nil? and @par.nil? then raise 'New file should be named on first save.' end
    
    if @par.kind_of? StringIO
      @par.write(@content_xml.to_s(indent: false))
    elsif @par.nil? or @par.kind_of? String
     
      if @par.kind_of? String   # the filename has changed 
        # first copy the original file to new location (or template if it is a new file)
        FileUtils.cp(@filename || File.dirname(__FILE__)+'/empty_file_template.ods', @par)
        @filename = @par
      end
      

      Zip::File.open(@filename) do |zip|
        # open manifest
        @manifest_xml = LibXML::XML::Document.io zip.get_input_stream('META-INF/manifest.xml')

        # save all pictures - iterate through sheets and pictures and check if they are saved and if not, save them
        @worksheets.each do |sheet|
          sheet.images.each do |image|
            # check if it is saved
            @ifname = image.internal_filename
            if @ifname.nil? or @ifname=='nic' or zip.find_entry(@ifname).nil?   
              # if it does not have name -> make up unused name 
              if @ifname.nil? or @ifname=='nic' 
                @ifname = image.internal_filename = Rspreadsheet::Tools.get_unused_filename(zip,'Pictures/',File.extname(image.original_filename))
              end
              raise 'Could not set up internal_filename correctly.' if @ifname.nil?
              
              # save it to zip file
              zip.add(@ifname, image.original_filename)
              
              # make sure it is in manifest
              if @manifest_xml.find("//manifest:file-entry[@manifest:full-path='#{@ifname}']").empty?
                node = Tools.prepare_ns_node('manifest','file-entry')
                Tools.set_ns_attribute(node,'manifest','full-path',@ifname)
                Tools.set_ns_attribute(node,'manifest','media-type',image.mime)
                @manifest_xml.find_first("//manifest:manifest") << node
              end  
            end
          end
        end

        zip.get_output_stream('content.xml') do |f|
          f.write @content_xml.to_s(:indent => false)
        end
        
        zip.get_output_stream('META-INF/manifest.xml') do |f|
          f.write @manifest_xml.to_s
        end
      end
    end
  end
  
  # Saves the worksheet to IO stream.
  def save_to_io(io = ::StringIO.new)
    ::Zip::OutputStream.write_buffer(io) do |output|
      ::Zip::File.open(TEMPLATE_FILE) do |input|
        input.
          select { |entry| entry.file? }.
          select { |entry| entry.name != CONTENT_FILE_NAME }.
          each do |entry|
            output.put_next_entry(entry.name)
            output.write(entry.get_input_stream.read)
          end
      end

      output.put_next_entry(CONTENT_FILE_NAME)
      output.write(@content_xml.to_s(indent: false))
    end
  end
  alias :to_io :save_to_io 
  
  private 
  CONTENT_FILE_NAME = 'content.xml'
  TEMPLATE_FILE = (File.dirname(__FILE__)+'/empty_file_template.ods').freeze
  def register_worksheet(worksheet)
    index = worksheets_count+1
    @worksheets[index-1]=worksheet
    @xmlnode << worksheet.xmlnode if worksheet.xmlnode.doc != @xmlnode.doc
  end
end
end
