require 'spec_helper'

$test_filename = './spec/testfile1.ods'

describe Rspreadsheet do

  it 'can open ods file' do
    book = Rspreadsheet.new($test_filename)
    book.worksheets[0].should_not == nil
    book.worksheets[0].class.should == Rspreadsheet::Worksheet
    s = book.worksheets[0]
    (1..10).each do |i|
      s[i-1,0].should === i
    end
    s[0,1].should === 'text'
    s[1,1].should === Date.new(2014,1,1)
  end
  
  it 'can open and save file, and saved file is the same as original' do
    pending
    tmp_filename = './tmp/testfile1.ods'    
    File.delete(tmp_filename) if File.exists?(tmp_filename)
    book = Rspreadsheet.new()
    book.save(tmp_filename)
    book1 = Rspreadsheet.new($test_filename)
    book2 = Rspreadsheet.new(tmp_filename)
    @sheet1 = book1.worksheets[0]
    @sheet2 = book2.worksheets[0]
  
    @sheet1.nonemptycells.each do |cell|
      @sheet2[cell.row,cell.col].should == cell.value
    end
  
  end
  
  it 'can create file' do
    book = Rspreadsheet.new
  end
  
  it 'can create new worksheet' do
    book = Rspreadsheet.new
    book.create_worksheet
  end

end

describe Rspreadsheet::Cell do
  before do 
    book1 = Rspreadsheet.new
    @sheet1 = book1.create_worksheet
    @sheet1[0,0] = 'text'
    book2 = Rspreadsheet.new($test_filename)
    @sheet2 = book2.worksheets[0]
  end
  it 'contains good row and col coordinates' do
    @cell = @sheet1.cells[1,3]
    @cell.row.should == 1
    @cell.col.should == 3
    @cell.coordinates.should == [1,3]
    
    @cell = @sheet2.cells[0,1]
    @cell.row.should == 0
    @cell.col.should == 1
    @cell.coordinates.should == [0,1]
  end
  it 'can be referenced by more vars and both are synchromized' do
    pending
    @cell = @sheet1.cells[0,0]
    @sheet1[0,0] = 'novinka'
    @cell.value.should == 'novinka'
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