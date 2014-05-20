require 'spec_helper'

describe Rspreadsheet do
  it 'can open ods file' do
    pending
    book = Rspreadsheet.open('./testfile1.ods')
    book.worksheets[0].cells[0,0].value.should === 1
  end
  
  it 'can create file' do
    book = Rspreadsheet.new
  end
  
  it 'can create new worksheet' do
    book = Rspreadsheet.new
    book.create_worksheet
  end

end

describe Rspreadsheet::Worksheet do
  before do
    book = Rspreadsheet.new
    @sheet = book.create_worksheet
  end

  it 'remembers the value stored to A1 cell' do
    @sheet[0,0].should == nil
    @sheet[0,0] = 'test text'
    @sheet[0,0].class.should == String
    @sheet[0,0].should == 'test text'
  end

  it 'value stored to A1 is accesible using different syntax' do
    @sheet[0,0] = 'test text'
    @sheet[0,0].should == 'test text'
    @sheet.cells[0,0].value.should == 'test text'
  end

  it 'makes Cell object accessible' do
    @sheet.cells[0,0].value = 'test text'
    @sheet.cells[0,0].class.should == Rspreadsheet::Cell
  end
  
  it 'has name, which can be changed and is remembered' do
    @sheet.name.should be(nil)
    @sheet.name = 'Icecream'
    @sheet.name.should == 'Icecream'
    @sheet.name = 'Cofee'
    @sheet.name.should == 'Cofee'    
  end
  
end