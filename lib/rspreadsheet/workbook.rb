require 'libxml'


module Rspreadsheet
class Workbook
  attr_reader :xmlnode # debug


  def xmldoc; @xmlnode.doc end

  #@!group Worskheets methods
  def create_worksheet_from_node(source_node)
    sheet = Worksheet.new(source_node)
    register_worksheet(sheet)
    return sheet
  end
  def create_worksheet(name = "Sheet#{worksheets_count+1}")
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
  alias :worksheet :worksheets
  alias :sheet :worksheets
  alias :sheets :worksheets
  def [](index_or_name); self.worksheets(index_or_name) end
  #@!group Loading and saving related methods
  def initialize
    @worksheets = []
  end

  # @param io [IO] reading which document will be loaded
  #
  # @return [self]
  def load(io)
    @content_xml = ::LibXML::XML::Document.io(io)
    @xmlnode     = @content_xml.find_first('//office:spreadsheet')
    @xmlnode.find('./table:table').each do |node|
      create_worksheet_from_node(node)
    end

    self
  end

  # @param io [IO] writing which document will be stored
  #
  # @return [self]
  def store(io)
    io.write(@content_xml.to_s(indent: false))

    self
  end

  private

  def register_worksheet(worksheet)
    index = worksheets_count+1
    @worksheets[index-1]=worksheet
    @xmlnode << worksheet.xmlnode if worksheet.xmlnode.doc != @xmlnode.doc
  end
end
end
