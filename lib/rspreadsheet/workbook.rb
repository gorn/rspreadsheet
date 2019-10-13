require 'zip'
require 'libxml'

module Rspreadsheet

class Workbook
  attr_reader :filename
  attr_reader :xmlnode # debug
  def xmldoc; @xmlnode.doc end

  #@!group Worksheet methods
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
  def worksheets_count
    @worksheets.length
  end
  alias :worksheet_count :worksheets_count
  alias :size :worksheets_count

  # @return [String] names of sheets in the workbook
  def worksheet_names
    @worksheets.collect{ |ws| ws.name }
  end

  # @param [Integer,String]
  # @return [Worksheet] worksheet with given index or name
  def worksheet(index_or_name)
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
  alias :sheet :worksheet

	# @param [Integer,String]
  # if index_or_name is provided, calls worksheet
	# @return Array of all worksheets in the document
	def worksheets(index_or_name=nil)
		case index_or_name
  		when nil
				@worksheets
			else
  			worksheet(index_or_name)
		end
	end

  alias :sheets :worksheets
  def [](index_or_name); self.worksheets(index_or_name) end
  #@!group Loading and saving related methods

  # @return Mime of the file
  def mime
    'application/vnd.oasis.opendocument.spreadsheet'.freeze
  end

  # @return [String] Prefered file extension
  def mime_preferred_extension
    'ods'.freeze
  end
  alias :mime_default_extension :mime_preferred_extension

  def initialize(afilename=nil)
    @worksheets=[]
    @filename = afilename
    @content_xml = Zip::File.open(@filename || TEMPLATE_FILE_NAME) do |zip|
      LibXML::XML::Document.io zip.get_input_stream(CONTENT_FILE_NAME)
    end
    @xmlnode = @content_xml.find_first('//office:spreadsheet')
    @xmlnode.find('./table:table').each do |node|
      create_worksheet_from_node(node)
    end
  end

  # @param [String] Optional new filename
  # Saves the worksheet. Optionally you can provide new filename or IO stream to which the file should be saved.
  def save(io=nil)
    case
      when @filename.nil? && io.nil?
        raise 'New file should be named on first save.'
      when @filename.kind_of?(String) && io.nil?
        Tools.output_to_zip_stream(@filename) do |input_and_output_zip|                  # open old file
          update_zip_manifest_and_content_xml(input_and_output_zip,input_and_output_zip) # input and output are identical
        end
      when (@filename.kind_of?(String) && (io.kind_of?(String) || io.kind_of?(File)))
        io = io.path if io.kind_of?(File)                                           # convert file to its filename
        FileUtils.cp(@filename , io)                                                # copy file externally
        @filename = io                                                              # remember new name
        save_to_io(nil)                                                             # continue modyfying file on spot
      when io.kind_of?(IO) || io.kind_of?(String) || io.kind_of?(StringIO)
        Tools.output_to_zip_stream(io) do |output_io|                               # open output stream of file
          write_ods_to_io(output_io)
        end
        io.rewind if io.kind_of?(StringIO)
      else raise 'Ivalid combinations of parameter types in save'
    end
  end
  alias :save_to_io  :save
  alias :save_as :save
  def to_io
    WorkbookIO.new(self)
  end

  def write_ods_to_io(io)
    if @filename.nil?
      Zip::File.open(TEMPLATE_FILE_NAME) do |empty_template_zip|         # open empty_template file
        copy_internally_without_content(empty_template_zip,io)           # copy empty_template internals
        update_zip_manifest_and_content_xml(empty_template_zip,io)           # update xmls + pictures
      end
    else
      Zip::File.open(@filename) do | old_zip |                           # open old file
        copy_internally_without_content(old_zip,io)                      # copy the old internals
        update_zip_manifest_and_content_xml(old_zip,io)                      # update xmls + pictures
      end
    end
  end

  def flat_format?; false end
  def normal_format?; true end

  private

  def update_zip_manifest_and_content_xml(input_zip,output_zip)
    update_manifest_xml(input_zip,output_zip)
    update_content_xml(output_zip)
  end

  def update_content_xml(zip)
    save_entry_to_zip(zip,CONTENT_FILE_NAME,@content_xml.to_s(indent: false))
  end

  def update_manifest_xml(input_zip,output_zip)
    # read manifest
    @manifest_xml = LibXML::XML::Document.io input_zip.get_input_stream(MANIFEST_FILE_NAME)
    modified = false
    # save all pictures - iterate through sheets and pictures and check if they are saved and if not, save them
    @worksheets.each do |sheet|
      sheet.images.each do |image|
        # check if it is saved
        @ifname = image.internal_filename
        if @ifname.nil? or input_zip.find_entry(@ifname).nil?
          # if it does not have name -> make up unused name
          if @ifname.nil?
            @ifname = image.internal_filename = Rspreadsheet::Tools.get_unused_filename(input_zip,'Pictures/',File.extname(image.original_filename))
          end
          raise 'Could not set up internal_filename correctly.' if @ifname.nil?
          raise 'This should not happen' if image.original_filename.nil?

          # save it to zip file
          save_entry_to_zip(output_zip, @ifname, File.open(image.original_filename,'r').read)

          # make sure it is in manifest
          if @manifest_xml.find("//manifest:file-entry[@manifest:full-path='#{@ifname}']").empty?
            node = Tools.prepare_ns_node('manifest','file-entry')
            Tools.set_ns_attribute(node,'manifest','full-path',@ifname)
            Tools.set_ns_attribute(node,'manifest','media-type',image.mime)
            @manifest_xml.find_first("//manifest:manifest") << node
            modified = true
          end
        end
      end
    end

    # write manifest if it was modified
    save_entry_to_zip(output_zip, MANIFEST_FILE_NAME,
                      @manifest_xml.to_s) if modified
  end

  def copy_internally_without_content(input_zip,output_zip)
    input_zip.each do |entry|
      next unless entry.file?
      next if entry.name == CONTENT_FILE_NAME
      save_entry_to_zip(output_zip, entry.name, entry.get_input_stream.read)
    end
  end

  def save_entry_to_zip(zip,internal_filename,contents)
    if zip.kind_of? Zip::File
      zip.get_output_stream(internal_filename) do  |f|
        f.write contents
      end
    else
      zip.put_next_entry(internal_filename)
      zip.write(contents)
    end
  end

  CONTENT_FILE_NAME = 'content.xml'
  MANIFEST_FILE_NAME = 'META-INF/manifest.xml'
  TEMPLATE_FILE_NAME = (File.dirname(__FILE__)+'/empty_file_template.ods').freeze
  def register_worksheet(worksheet)
    index = worksheets_count+1
    @worksheets[index-1]=worksheet
    @xmlnode << worksheet.xmlnode if worksheet.xmlnode.doc != @xmlnode.doc
  end

end

class WorkbookFlat < Workbook
  def initialize(afilename=nil)
    @worksheets=[]
    @filename = afilename
    @xml_doc = LibXML::XML::Document.file(@filename || FLAT_TEMPLATE_FILE_NAME)
    @xmlnode = @xml_doc.find_first('//office:spreadsheet')
    @xmlnode.find('./table:table').each do |node|
      create_worksheet_from_node(node)
    end
  end

  def save(io=nil)
    case
      when @filename.nil? && io.nil?
        raise 'New file should be named on first save, please provide filename (or IO).'
      when @filename.kind_of?(String) && io.nil?
        @xml_doc.save(@filename)
      when (@filename.kind_of?(String) && (io.kind_of?(String) || io.kind_of?(File)))
        @filename = (io.kind_of?(File)) ? io.path : io
        @xml_doc.save(@filename)
      when io.kind_of?(IO) || io.kind_of?(String) || io.kind_of?(StringIO)
        IO.write(io,@xml_doc.to_s)
        io.rewind if io.kind_of?(StringIO)
      else raise 'Invalid combinations of parameter types in save'
    end
  end
  alias :save_to_io  :save
  alias :save_as :save

  def flat_format?; true end
  def normal_format?; false end

  private
  FLAT_TEMPLATE_FILE_NAME = (File.dirname(__FILE__)+'/empty_file_template.fods').freeze

end

class WorkbookIO
  def initialize(workbook)
    @workbook = workbook
  end
  def read
    buffer.string
  end
  private
  def buffer
    Zip::OutputStream.write_buffer do |io|
      @workbook.write_ods_to_io(io)
    end
  end
end

end
