require 'spec_helper'

describe Rspreadsheet::Worksheet do
  before do 
    @sheet = Rspreadsheet.new($test_filename).worksheets[1]
  end
  it 'contains nonempty xml in rows for testfile' do
    @sheet.rows(1).xmlnode.elements.size.should be >1
  end
  it 'uses detach_my_subnode_respect_repeated well' do
    @sheet.detach_my_subnode_respect_repeated(50, {:xml_items_node_name => 'table-row', :xml_repeated_attribute => 'number-rows-repeated'})
    @sheet.rows(50).detach_my_subnode_respect_repeated(12, {:xml_items_node_name => 'table-cell', :xml_repeated_attribute => 'number-columns-repeated'})
  end
end

describe Rspreadsheet::Worksheet do
  before do
    book = Rspreadsheet.new
    @sheet = book.create_worksheet
  end
  it 'freshly created has correctly namespaced xmlnode' do
    @xmlnode = @sheet.xmlnode
    @xmlnode.namespaces.to_a.size.should >5
    @xmlnode.namespaces.find_by_prefix('office').should_not be_nil
    @xmlnode.namespaces.find_by_prefix('table').should_not be_nil
    @xmlnode.namespaces.namespace.should_not be_nil
    @xmlnode.namespaces.namespace.prefix.should == 'table'
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
  it 'out of range indexes return nil value' do
    @sheet[-1,-1].should == nil
    @sheet[0,0].should == nil
    @sheet[999,999].should == nil
  end
  it 'returns nil with negative index' do
    @sheet.rows(-1).should == nil
  end
end 