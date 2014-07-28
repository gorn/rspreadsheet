require 'spec_helper'

describe Rspreadsheet::Workbook do
  it 'has correct number of sheets' do
    book = Rspreadsheet.new($test_filename)
    book.worksheets_count.should == 1
    book.worksheets[0].should be_nil
    book.worksheets[1].should be_kind_of(Rspreadsheet::Worksheet)
    book.worksheets[2].should be_nil
    book.worksheets[nil].should be_nil
  end
end
