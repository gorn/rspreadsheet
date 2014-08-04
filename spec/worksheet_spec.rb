require 'spec_helper'

describe Rspreadsheet::Worksheet do
  before do 
    @sheet = Rspreadsheet.new($test_filename).worksheets[1]
  end
  it 'contains nonempty xml in rows for testfile' do
    @sheet.rows(1).xmlnode.elements.size.should be >1
  end
  it 'freshly created has correctly namespaced xmlnode' do
    @spreadsheet = Rspreadsheet.new
    @spreadsheet.create_worksheet
    @xmlnode = @spreadsheet.worksheets[1].xmlnode
    @xmlnode.namespaces.to_a.size.should >5
    @xmlnode.namespaces.find_by_prefix('office').should_not be_nil
    @xmlnode.namespaces.find_by_prefix('table').should_not be_nil
    @xmlnode.namespaces.namespace.should_not be_nil
    @xmlnode.namespaces.namespace.prefix.should == 'table'
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