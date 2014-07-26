require 'spec_helper'

$test_filename = './spec/testfile1.ods'

describe Rspreadsheet::Tools do
  it '=Converts correctly cell adresses' do
    Rspreadsheet::Tools.convert_cell_address('A1') [0].should == 1
    Rspreadsheet::Tools.convert_cell_address('A1')[1].should == 1
    Rspreadsheet::Tools.convert_cell_address('C5')[0].should == 5
    Rspreadsheet::Tools.convert_cell_address('C5')[1].should == 3
  end
end

describe Rspreadsheet::Tools::SparseRepeatedArray do
  before do
#     @a = Rspreadsheet::Tools::SparseRepeatedArray.new()
  end
#   Float::INFINITY
  its 'set values can be read back unchanged' do
    skip
    @a[3..5]='test'
    @a[4].should == 'test'
  end
end

describe Rspreadsheet do
  it 'can open ods testfile and reads its content correctly' do
    book = Rspreadsheet.new($test_filename)
    book.worksheets[0].should_not == nil
    book.worksheets[0].class.should == Rspreadsheet::Worksheet
    s = book.worksheets[0]
    (1..10).each do |i|
      s[i,1].should === i
    end
    s[1,2].should === 'text'
    s[2,2].should === Date.new(2014,1,1)
  end
  it 'can open and save file, and saved file has same cells as original' do
    tmp_filename = '/tmp/testfile1.ods'        # first delete temp file
    File.delete(tmp_filename) if File.exists?(tmp_filename)
    book = Rspreadsheet.new($test_filename)    # than open test file
    book.save(tmp_filename)                    # and save it as temp file
    
    book1 = Rspreadsheet.new($test_filename)   # now open both again
    book2 = Rspreadsheet.new(tmp_filename)
    @sheet1 = book1.worksheets[0]
    @sheet2 = book2.worksheets[0]
    
    @sheet1.nonemptycells.each do |cell|       # and test identity
      @sheet2[cell.row,cell.col].should == cell.value
    end
  end
  it 'can open and save file, and saved file is exactly same as original' do
    tmp_filename = '/tmp/testfile1.ods'        # first delete temp file
    File.delete(tmp_filename) if File.exists?(tmp_filename)
    book = Rspreadsheet.new($test_filename)    # than open test file
    book.save(tmp_filename)                    # and save it as temp file
    
    # now compare them
    @content_xml1 = Zip::File.open($test_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_xml2 = Zip::File.open(tmp_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_xml1.root.equals?(@content_xml2.root).should == true
  end
  it 'when open and save file modified, than the file is different' do
    tmp_filename = '/tmp/testfile1.ods'        # first delete temp file
    File.delete(tmp_filename) if File.exists?(tmp_filename)
    book = Rspreadsheet.new($test_filename)    # than open test file
    book.worksheets[0][1,1].should_not == 'xyzxyz'
    book.worksheets[0][1,1]='xyzxyz'
    book.worksheets[0][1,1].should == 'xyzxyz'
    book.save(tmp_filename)                    # and save it as temp file
    
    # now compare them
    @content_doc1 = Zip::File.open($test_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_doc2 = Zip::File.open(tmp_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_doc1.eql?(@content_doc2).should == false
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
    @cell = @sheet1.cells(1,3)
    @cell.row.should == 1
    @cell.col.should == 3
    @cell.coordinates.should == [1,3]
    
    @cell = @sheet2.cells(7,2)
    @cell.row.should == 7
    @cell.col.should == 2
    @cell.coordinates.should == [7,2]
  end
  it 'can be referenced by more vars and both are synchronized' do
    @cell = @sheet1.cells(1,1)
    @sheet1[1,1] = 'novinka'
    @cell.value.should == 'novinka'
  end
  it 'can be modified by more ways and all are identical' do
    @cell = @sheet1.cells(2,2)
    @sheet1[2,2] = 'zaprve'
    @cell.value.should == 'zaprve'
    @sheet1.cells(2,2).value = 'zadruhe'
    @cell.value.should == 'zadruhe'
    @sheet1.B2 = 'zatreti'
    @cell.value.should == 'zatreti'
  end
  it 'can include links' do
    @sheet2.A12.should == '[http://example.org/]'
  end
  it 'contains good row and col coordinates even after table:number-columns-repeated cells' do
    skip
    @cell = @sheet2.cells(13,5)
    @cell.value.should == 'afterrepeated'
    @cell.row.should == 13
    @cell.col.should == 5
  end
end

describe Rspreadsheet::Worksheet do
  before do
    book = Rspreadsheet.new
    @sheet = book.create_worksheet
  end
  it 'remembers the value stored to A1 cell' do
    @sheet[1,1].should == nil
    @sheet[1,1] = 'test text'
    @sheet[1,1].class.should == String
    @sheet[1,1].should == 'test text'
  end
  it 'value stored to A1 is accesible using different syntax' do
    @sheet[1,1] = 'test text'
    @sheet[1,1].should == 'test text'
    @sheet.cells(1,1).value.should == 'test text'
  end
  it 'makes Cell object accessible' do
    @sheet.cells(1,1).value = 'test text'
    @sheet.cells(1,1).class.should == Rspreadsheet::Cell
  end
  it 'has name, which can be changed and is remembered' do
    @sheet.name.should be(nil)
    @sheet.name = 'Icecream'
    @sheet.name.should == 'Icecream'
    @sheet.name = 'Cofee'
    @sheet.name.should == 'Cofee'    
  end
end

describe Rspreadsheet::Row do
  before do 
    book1 = Rspreadsheet.new
    @sheet1 = book1.create_worksheet
  end
  it 'allows access to cells in a row' do
    (2..5).each { |i| @sheet1[7,i] = i }
    (2..5).each { |i| 
      a = @sheet1.rows(7)
      c = a.cells(i)
      c.value.should == i 
    }
  end
end
