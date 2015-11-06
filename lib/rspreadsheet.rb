require 'rspreadsheet/version'
require "rspreadsheet/document"
require 'rspreadsheet/workbook'
require 'rspreadsheet/worksheet'
require 'helpers/class_extensions'
require 'helpers/configuration'


# @example
#
#   # Create new document
#   document = ::Rspreadsheet.create_document
#   workbook = document.workbook
#
#   # Open existead document
#   document = ::Rspreadsheet.open_document(path_to_the_spreadsheet_file)
#   workbook = document.workbook
#
#   # Save document to the file.
#   ::Rspreadsheet.save_document(document, path_to_a_new_spreadsheet_file)
module Rspreadsheet
  extend Configuration
  define_setting :raise_on_negative_coordinates, true


  class << self
    # @return [::Rspreadsheet::Document]
    def create_document
      document          = ::Rspreadsheet::Document.new
      document.workbook = ::Rspreadsheet::Workbook.new

      document.build_zip_input_stream_from_file(::Rspreadsheet::Document::TEMPLATE_FILE) do |stream|
        document.workbook.load(stream)
      end

      document
    end

    # @param file_path [::Rspreadsheet::Document] file path to the document
    #   file which should be opened
    #
    # @return [::Rspreadsheet::Document]
    def open_document(file_path)
      document          = ::Rspreadsheet::Document.new
      document.workbook = ::Rspreadsheet::Workbook.new

      document.build_zip_input_stream_from_file(file_path) do |stream|
        document.workbook.load(stream)
      end

      document
    end

    # Saves the worksheet. Optionally you can provide new filename.
    #
    # @param document  [::Rspreadsheet::Document]
    # @param file_path [String, Pathname] (nil) file path to file where document
    #   will be saved
    #
    # @return [::Rspreadsheet::Document]
    #
    # @raise [::RuntimeError] when given document has no file path and path to
    #   save file has not been given
    # @raise [::RuntimeError] when by given file_path or document's file_path
    #   file already exists.
    def save_document(document, file_path = nil)
      file_save_path = file_path || document.file_path
      unless file_save_path
        raise "New file should be named on first save."
      end

      if ::File.exists?(file_save_path)
        raise "File at path '#{file_save_path}' already exists."
      end

      ::File.new(file_save_path) do |ods_file_io|
        document.save(ods_file_io)
      end
    end
  end
end
