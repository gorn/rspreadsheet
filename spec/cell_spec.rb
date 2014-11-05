require 'spec_helper'
 
describe Rspreadsheet::Cell do
  before do 
    book1 = Rspreadsheet.new
    @sheet1 = book1.create_worksheet
    book2 = Rspreadsheet.new($test_filename)
    @sheet2 = book2.worksheets[1]
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
    @cell = @sheet2.cells(13,5)
    @cell.value.should == 'afterrepeated'
    @cell.row.should == 13
    @cell.col.should == 5
  end
  it 'does not accept negative and zero coordinates' do
    @sheet2.cells(0,5).should be(nil)
    @sheet2.cells(2,-5).should be(nil)
    @sheet2.cells(-2,-5).should be(nil)
  end
  it 'has nonempty parents' do
    @cell = @sheet2.cells(13,5)
    @cell.parent_row.should_not be_nil
    @cell.worksheet.should_not be_nil

    @cell = @sheet1.cells(2,2)
    @cell.parent_row.should_not be_nil
    @cell.worksheet.should_not be_nil
  end
  it 'handles relative correctly' do
    @sheet2.cells(3,3).relative(-1,+2).coordinates.should == [2,5]
    @sheet2.cells(3,3).relative(0,0).coordinates.should == [3,3]
  end
  it 'is automatically "unrepeated" on value assignement' do
    @cell = @sheet2.cells(13,2)
    @cell.is_repeated?.should == true
#     binding.pry
#     @cell.value = 'cokoli'
#     @cell.is_repeated?.should == false
#     @cell.value.should == 'cokoli'
#     @sheet2.cells(13,1).should_not == 'cokoli'
#     @sheet2.cells(13,3).should_not == 'cokoli'
#     @sheet2.cells(13,4).should_not == 'cokoli'
  end
end





