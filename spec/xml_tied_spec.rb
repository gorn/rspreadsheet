require 'spec_helper'
require 'rspreadsheet/xml_tied_array'

describe Rspreadsheet::XMLTiedArray do
	before do
		@book = Rspreadsheet.new
		@sheet = @book.create_worksheet
    class TestXMLTiedArray 
      include Rspreadsheet::XMLTiedArray
    end
	end
  it 'method subitems does not accept 2 parameteres' do
    expect {@sheet.subitems(1,2)}.to raise_error ArgumentError
  end
  it 'does not have xmlnode method by default' do
    tx = TestXMLTiedArray.new
    expect {tx.xmlnode}.to raise_error
  end
  
  it 'raises when prepare_empty_xmlnode fails in insert_new_empty_subnode_before' do
    class TestXMLTiedArray 
      def subitem_xml_options; {} end
      def xmlnode; nil end
    end
    
    tx = TestXMLTiedArray.new
    expect {tx.insert_new_empty_subnode_before(0)}.to raise_error IndexError
    expect {tx.insert_new_empty_subnode_before(1)}.to raise_error(/create empty xmlnode/)
  end
end
