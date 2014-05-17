require 'spec_helper'

describe Rspreadsheet do
  it 'can open ods file' do
    pending
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
    @sheet.cell(0,0).value = 'test text'
    @sheet.cell(0,0).value.should == 'test text'
  end

  it 'remembers the value stored to A1 using array syntax' do
    pending
    @sheet.cell(0,0).should_be_nil
    @sheet.cells[0,0].should_be_nil
 
    @sheet.cells[0,0] = 'test text'
    @sheet.cells[0,0].should == 'test text'
    @sheet.cell(0,0).should == 'test text'
  end

  
  it 'has name, which can be changed and is remembered' do
    @sheet.name.should be(nil)
    @sheet.name = 'Icecream'
    @sheet.name.should == 'Icecream'
    @sheet.name = 'Cofee'
    @sheet.name.should == 'Cofee'    
  end
  
end