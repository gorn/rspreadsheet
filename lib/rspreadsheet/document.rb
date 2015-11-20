require "zip"


# This is an ODS container which is used to serialize document.
class ::Rspreadsheet::Document
  # @!attribute MIME_TYPE [r]
  #   @return [String]
  MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet".freeze

  # @!attribute FILE_EXTENSION [r]
  #   @return [String]
  FILE_EXTENSION = "ods".freeze

  # @!attribute TEMPLATE_FILE
  #   @return [Pathname]
  TEMPLATE_FILE = Pathname.new(__dir__).join("empty_file_template.ods").freeze

  # @!attribute CONTENT_FILE_NAME
  #   @return [String]
  CONTENT_FILE_NAME = "content.xml".freeze


  # @!attribute file_path
  #   @return [String, Path]
  attr_accessor \
    :file_path,
    :workbook,
    :mime_type,
    :file_extension


  # @return [void]
  def initialize
    @file_path = @workbook = nil

    @mime_type      = MIME_TYPE
    @file_extension = FILE_EXTENSION
  end

  # @return [void]
  def build_zip_input_stream_from_file(file_path)
    ::Zip::File.open(file_path) do |zip|
      yield zip.get_input_stream(CONTENT_FILE_NAME)
    end
  end

  # @param io [IO] (StringIO)
  #
  # @return [IO]
  def save(io = ::StringIO.new)
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
      @workbook.store(output)
    end
  end
end
