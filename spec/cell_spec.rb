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
    @cell.rowi.should == 1
    @cell.coli.should == 3
    @cell.coordinates.should == [1,3]
    
    @cell = @sheet2.cells(7,2)
    @cell.rowi.should == 7
    @cell.coli.should == 2
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
    @cell.rowi.should == 13
    @cell.coli.should == 5
  end
  it 'reports good range of coordinates for repeated cells' do
    @cell = @sheet2.cells(13,2)
    @cell.range.should == (1..4)
    @cell.mode.should == :repeated
  end
  it 'does not accept negative and zero coordinates' do
    @sheet2.cells(0,5).should be(nil)
    @sheet2.cells(2,-5).should be(nil)
    @sheet2.cells(-2,-5).should be(nil)
  end
  it 'has nonempty parents' do
    @cell = @sheet2.cells(13,5)
    @cell.row.should_not be_nil
    @cell.worksheet.should_not be_nil

    @cell = @sheet1.cells(2,2)
    @cell.row.should_not be_nil
    @cell.worksheet.should_not be_nil
  end
  it 'handles relative correctly' do
    @sheet2.cells(3,3).relative(-1,+2).coordinates.should == [2,5]
    @sheet2.cells(3,3).relative(0,0).coordinates.should == [3,3]
  end
  it 'is automatically "unrepeated" on value assignement' do
    @cell = @sheet2.cells(13,2)
    @cell.is_repeated?.should == true
    @cell.value = 'cokoli'
    @cell.is_repeated?.should == false
    @cell.value.should == 'cokoli'
    @sheet2.cells(13,1).should_not == 'cokoli'
    @sheet2.cells(13,3).should_not == 'cokoli'
    @sheet2.cells(13,4).should_not == 'cokoli'
  end
  it 'returns type for the cell' do
    book = Rspreadsheet.new($test_filename)
    s = book.worksheets[1]
    s.cells(1,2).type.should === :string
    s.cells(2,2).type.should === :date
    s.cells(3,1).type.should === :float
    s.cells(3,2).type.should === :percentage
    s.cells(4,2).type.should === :string
    s.cells(200,200).type.should === :unassigned
  end
  it 'is the same object no matter how you access it' do
    @cell1 = @sheet2.cells(5,5)
    @cell2 = @sheet2.rows(5).cells(5)
    @cell1.should equal(@cell2)
  end
  it 'splits correctly cells if written in the middle of repeated group' do
    @cell = @sheet2.cells(4,6)
    @cell.range.should == (4..7)
    @cell.value.should == 7
    
    @cell.value = 'nebesa'
    @cell.range.should == (6..6)
    @cell.value.should == 'nebesa'
    
    @cellA = @sheet2.cells(4,5)
    @cellA.range.should == (4..5)
    @cellA.value.should == 7
    
    @cellB = @sheet2.cells(4,7)
    @cellB.range.should == (7..7)
    @cellB.value.should == 7
  end
  it 'inserts correctly cell in the middle of repeated group' do
    @cell = @sheet2.cells(4,6)
    @cell.range.should == (4..7)
    @cell.value.should == 7
    
    @sheet2.insert_cell_before(4,6)
    
    @cellA = @sheet2.cells(4,5)
    @cellA.range.should == (4..5)
    @cellA.value.should == 7
    
    @cellB = @sheet2.cells(4,7)
    @cellB.range.should == (7..8)
    @cellB.value.should == 7
    
    @cell = @sheet2.cells(16,4)
    @cell.range.should == (1..7)
    @cell.value.should == ""
    
    @sheet2.rows(15).range.should == (14..18)
    @sheet2.rows(16).range.should == (14..18)
    @sheet2.rows(17).range.should == (14..18)
    @sheet2.insert_cell_before(16,3)
    @sheet2.cells(16,3).value = 'baf'
    @sheet2.cells(17,3).value.should_not == 'baf'
    @sheet2.rows(15).range.should == (14..15)
    @sheet2.rows(16).range.should == (16..16)
    @sheet2.rows(17).range.should == (17..18)
    
    @cellA = @sheet2.cells(16,1)
    @cellA.range.should == (1..2)
    @cellA.value.should == ""
    
    @cellB = @sheet2.cells(16,5)
    @cellB.range.should == (4..8)
    @cellB.value.should == ""

  end
  it 'inserted has correct class' do # based on real error
    @sheet2.insert_cell_before(1,1)
    @sheet2.rows(1).cells(1).should be_kind_of(Rspreadsheet::Cell)
  end
end





