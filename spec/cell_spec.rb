require 'spec_helper'

describe Rspreadsheet::Cell do
  before do 
    book1 = Rspreadsheet.new
    @sheet1 = book1.create_worksheet
    book2 = Rspreadsheet.new($test_filename)
    @sheet2 = book2.worksheets(1)
  end
  it 'contains good row and col coordinates' do
    @cell = @sheet1.cell(1,3)
    @cell.rowi.should == 1
    @cell.coli.should == 3
    @cell.coordinates.should == [1,3]
    @cell = @sheet2.cell(7,2)
    @cell.rowi.should == 7
    @cell.coli.should == 2
    @cell.coordinates.should == [7,2]
  end
  it 'can be referenced by more vars and both are synchronized' do
    @cell = @sheet1.cell(1,1)
    @sheet1[1,1] = 'novinka'
    @cell.value.should == 'novinka'
  end
  it 'can be modified by more ways and all are identical' do
    @cell = @sheet1.cell(2,2)
    @sheet1[2,2] = 'zaprve'
    @cell.value.should == 'zaprve'
    @sheet1.cell(2,2).value = 'zadruhe'
    @cell.value.should == 'zadruhe'
    @sheet1.B2 = 'zatreti'
    @cell.value.should == 'zatreti'
    @sheet1.row(2).cell(2).value = 'zactvrte'
    @cell.value.should == 'zactvrte'
    @sheet1.rows[1].cell(2).value = 'zactvrte s arrayem'
    @cell.value.should == 'zactvrte s arrayem'
    @sheet1.row(2)[2] = 'zapate'
    @cell.value.should == 'zapate'
    @sheet1.rows[1][2] = 'zaseste'  ## tohle je divoka moznost
    @cell.value.should == 'zaseste'
    
    @sheet1.A11 = 'dalsi test'
  end
  it 'can include links' do
    @sheet2.A12.should == '[http://example.org/]'
    @sheet2.cell(12,2).valuexmlfindall('.//text:a').size.should eq 0
    @sheet2.cell(12,1).valuexmlfindall('.//text:a').size.should eq 1
    @sheet2.cell(12,1).valuexmlfindfirst('.//text:a').attributes['href'].should eq 'http://example.org/'
  end
  it 'contains good row and col coordinates even after table:number-columns-repeated cells' do
    @cell = @sheet2.cell(13,5)
    @cell.value.should == 'afterrepeated'
    @cell.rowi.should == 13
    @cell.coli.should == 5
  end
  it 'reports good range of coordinates for repeated cells' do
    @cell = @sheet2.cell(13,2)
    @cell.range.should == (1..4)
    @cell.mode.should == :repeated
  end
  it 'returns nil on negative and zero cell indexes or raises exception depending on configuration' do
    pom = Rspreadsheet.raise_on_negative_coordinates
    # default is to raise error
    expect {@sheet2.cell(0,5) }.to raise_error
  
    # return nil if configured to do so
    Rspreadsheet.raise_on_negative_coordinates = false
    @sheet2.cell(0,5).should be_nil
    @sheet2.cell(2,-5).should be(nil)
    @sheet2.cell(-2,-5).should be(nil)
    
    Rspreadsheet.raise_on_negative_coordinates = pom  # reset the setting back
  end
  it 'has nonempty parents' do
    @cell = @sheet2.cell(13,5)
    @cell.row.should_not be_nil
    @cell.worksheet.should_not be_nil

    @cell = @sheet1.cell(2,2)
    @cell.row.should_not be_nil
    @cell.worksheet.should_not be_nil
  end
  it 'handles relative correctly' do
    @sheet2.cell(3,3).relative(-1,+2).coordinates.should == [2,5]
    @sheet2.cell(3,3).relative(0,0).coordinates.should == [3,3]
  end
  it 'is automatically "unrepeated" on value assignement' do
    @cell = @sheet2.cell(13,2)
    @cell.is_repeated?.should == true
    @cell.value = 'cokoli'
    @cell.is_repeated?.should == false
    @cell.value.should == 'cokoli'
    @sheet2.cell(13,1).should_not == 'cokoli'
    @sheet2.cell(13,3).should_not == 'cokoli'
    @sheet2.cell(13,4).should_not == 'cokoli'
  end
  it 'returns correct type for the cell' do
    @sheet2.cell(1,2).type.should eq :string
    @sheet2.cell(2,2).type.should eq :datetime
    @sheet2.cell(3,1).type.should eq :float
    @sheet2.cell(3,2).type.should eq :percentage
    @sheet2.cell(4,2).type.should eq :string
    @sheet2.cell('B22').type.should eq :currency
    @sheet2.cell('B23').type.should eq :currency
    @sheet2.cell(200,200).type.should eq :unassigned
  end
  it 'returns value of correct type' do
    @sheet2[1,2].should be_kind_of(String)
    @sheet2[2,2].should be_kind_of(DateTime)  # maybe date
    @sheet2[3,1].should be_kind_of(Float)
    @sheet2[3,2].should be_kind_of(Float)
    @sheet2.cell(3,2).type.should eq :percentage
    @sheet2.cell(3,2).guess_cell_type.should eq :percentage
    @sheet2.cell(3,2).guess_cell_type(1).should eq :percentage
    @sheet2[3,2]=0.1
    @sheet2.cell(3,2).type.should eq :percentage
    @sheet2[4,2].should be_kind_of(String)
  end
  it 'is the same object no matter how you access it' do
    @cell1 = @sheet2.cell(5,5)
    @cell2 = @sheet2.rows(5).cell(5)
    @cell1.should equal(@cell2)
  end
  it 'splits correctly cells if written in the middle of repeated group' do
    @cell = @sheet2.cell(4,6)
    @cell.range.should == (4..7)
    @cell.value.should == 7
    
    @cell.value = 'nebesa'
    @cell.range.should == (6..6)
    @cell.value.should == 'nebesa'
    
    @cellA = @sheet2.cell(4,5)
    @cellA.range.should == (4..5)
    @cellA.value.should == 7
    
    @cellB = @sheet2.cell(4,7)
    @cellB.range.should == (7..7)
    @cellB.value.should == 7
  end
  it 'inserts correctly cell in the middle of repeated group' do
    @cell = @sheet2.cell(4,6)
    @cell.range.should == (4..7)
    @cell.value.should == 7
    @cell.coli.should == 6
    
    @sheet2.insert_cell_before(4,6)
    @cell.coli.should == 7
    
    @cellA = @sheet2.cell(4,5)
    @cellA.range.should == (4..5)
    @cellA.value.should == 7
    
    @cellB = @sheet2.cell(4,7)
    @cellB.range.should == (7..8)
    @cellB.value.should == 7
    
    @cell = @sheet2.cell(16,4)
    @cell.range.should == (1..7)
    @cell.value.should == nil
    
    @sheet2.rows(15).range.should == (14..18)
    @sheet2.rows(16).range.should == (14..18)
    @sheet2.rows(17).range.should == (14..18)
    @sheet2.insert_cell_before(16,3)
    @sheet2.cell(16,3).value = 'baf'
    @sheet2.cell(17,3).value.should_not == 'baf'
    @sheet2.rows(15).range.should == (14..15)
    @sheet2.rows(16).range.should == (16..16)
    @sheet2.rows(17).range.should == (17..18)
    
    @cellA = @sheet2.cell(16,1)
    @cellA.range.should == (1..2)
    @cellA.value.should be_nil
    
    @cellB = @sheet2.cell(16,5)
    @cellB.range.should == (4..8)
    @cellB.value.should be_nil

  end
  it 'inserted has correct class' do # based on real error
    @sheet2.insert_cell_before(1,1)
    @sheet2.rows(1).cell(1).should be_kind_of(Rspreadsheet::Cell)
  end
  it 'method cells without arguments returns array of cells' do
    @a = @sheet2.rows(1).cells
    @a.should be_kind_of(Array)
    @a.each { |item| item.should be_kind_of(Rspreadsheet::Cell)}
  
  end
  it 'changes coordinates when row inserted above' do
    @sheet1.cell(2,2).detach
    @cell = @sheet1.cell(2,2)
    @cell.rowi.should == 2
    @sheet1.add_row_above(1)
    @cell.rowi.should == 3
  end
  it 'switches to invalid_reference cell when deleted' do
    @sheet1[2,5] = 'nejaka data'
    @cell = @sheet1.cell(2,2)
    @cell.value = 'data'
    @cell.invalid_reference?.should be false
    @cell.delete
    @cell.invalid_reference?.should be true
    expect { @cell.rowi }.to raise_error
    expect { @cell.address }.to raise_error
    
    @sheet1.cell(2,2).type.should == :empty
    @sheet1.cell(3,2).type.should == :unassigned
  end
  it 'switches to invalid_reference cell when its row is deleted' do
    @cell = @sheet1.cell(6,2)
    @cell.value = 'data'
    @cell.rowi.should == 6
    @sheet1.rows(6).delete
    expect { @cell.rowi }.to raise_error
    @cell.invalid_reference?.should be true
  end
  it 'has inspect method returning something good' do
    @cell = @sheet1.cell(6,2)
    @cell.value = 'abcde'
    expect(@cell.inspect).to include('abcde','::Cell','6','2','row')
  end
  it 'stores date correctly' do
    @cell = @sheet1.cell(1,1)
    @cell.value= Date.parse('2014-01-02')
    @cell.value.year.should eq 2014
    @cell.value.month.should eq 1
    @cell.value.day.should eq 2
  end
  it 'stores time correctly' do
    @cell = @sheet1.cell(1,1)
    @cell.value= Time.parse('2:42 pm')
    @cell.value.hour.should eq 14
    @cell.value.min.should eq 42
    @cell.value.sec.should eq 0
  end
  it 'parse_time_value converts correcty different time values' do
    dyear = 1899; dmonth = 12; dday = 30
    Rspreadsheet::Cell.parse_time_value('PT923451H33M00S').should == Time.new(2005,5,5,3,33,00,0)
    Rspreadsheet::Cell.parse_time_value('PT1H33M00S').should == Time.new(dyear,dmonth,dday,1,33,00,0)
  end
  it 'handles time of day correctly on assignement' do
    @sheet1.A11 = Rspreadsheet::Tools.new_time_value(2,13,27)
    @sheet1.A12 = @sheet1.A11
    @sheet1.A12.should == @sheet1.A11
    @sheet1.cell('A12').type.should == :time
  end
  it 'can read various types of times' do
    expect {@cell = @sheet2.cell('D22'); @cell.value }.not_to raise_error
    @cell.value.hour.should == 2
    @cell.value.min.should == 22
    expect {@cell = @sheet2.cell('D23'); @cell.value }.not_to raise_error
    @cell.value.should == Time.new(2005,5,5,3,33,0,0)
  end
  it 'handles dates and datetimes correctly on assignement' do
    @sheet1.A11 = DateTime.new(2011,1,1,2,13,27,0)
    @sheet1.A12 = @sheet1.A11
    @sheet1.A12.should == @sheet1.A11
    @sheet1.cell('A12').type.should == :datetime
    @sheet1.A11 = Date.new(2012,2,2)
    @sheet1.A12 = @sheet1.A11
    @sheet1.A12.should == @sheet1.A11
    @sheet1.cell('A12').type.should == :datetime
  end
  it 'can read various types of dates' do
    expect {@cell = @sheet2.cell('F22'); @cell.value }.not_to raise_error
    @cell.value.should == DateTime.new(2006,6,6,3,44,21,0)
  end
  it 'can be addressed by even more ways and all are identical' do
    @cell = @sheet1.cell(2,2)
    @sheet1.cell('B2').value = 'zaseste'
    @sheet1.cell('B2').value.should == 'zaseste'
    @cell.value.should == 'zaseste'
    @sheet1.cell(2,'B').value.should == 'zaseste'
    @sheet1.cell(2,'B').value = 'zasedme'
    @cell.value.should == 'zasedme'
    @sheet1['B2'].should == 'zasedme'
    @sheet1['B2'] = 'zaosme'
    @cell.value.should == 'zaosme'
    
    @sheet2.cell('F2').should be @sheet2.cell(2,6)
    @sheet2.cell('BA177').should be @sheet2.cell(177,53)
    @sheet2.cell('ADA2').should be @sheet2.cell(2,781)
  end
  it 'remembers formula when set' do
    @cell = @sheet1.cell(1,1)
    @cell.formula.should be_nil
    @cell.formula='=1+5'
    @cell.formula.should eq '=1+5'
  end
  it 'unsets cell type when formula set - we can not guess it correctly' do
    @cell = @sheet1.cell(1,1)
    @cell.value = 'ahoj'
    @cell.type.should eq :string
    @cell.formula='=1+5'
    typ = @cell.xmlnode.nil? ? 'N/A' : @cell.xmlnode.attributes['value-type']
    @cell.type.should_not eq :string
    @cell.type.should eq :empty
  end
  it 'wipes out formula after assiging value' do
    @cell = @sheet1.cell(1,1)
    @cell.formula='=1+5'
    @cell.formula.should_not be_nil
    @cell.value = 'baf'
    @cell.type.should eq :string
    @cell.formula.should be_nil
  end
  it 'works well with currency types' do
    @usdcell = @sheet2.cell('B22')
    @usdcell.type.should eq :currency
    @usdcell.value.should == -147984.84
    @usdcell.format.currency.should == 'USD'
    
    @czkcell = @sheet2.cell('B23')
    @czkcell.value.should == 344.to_d
    @czkcell.format.currency.should == 'CZK'
    
    @czkcell.value = 200.to_d
    @czkcell.value.should == 200.to_d
    @czkcell.format.currency.should == 'CZK'
  end
  it 'preserves currency format when float is assingned to it' do
    @cell = @sheet2.cell('B23')
    @cell.type.should eq :currency
    @cell.format.currency.should == 'CZK'
    
    @cell.value = 666.66.to_d
    @cell.value.should == 666.66.to_d
    @cell.type.should eq :currency
    @cell.format.currency.should == 'CZK'
  end
  
  it 'gracefully accepts nil in assignement' do
    expect {
      @sheet2.cell('B1').value = nil
      @sheet2.cell('B2').value = nil
      @sheet2.cell('B3').value = nil
      @sheet2.cell('B22').value = nil
      @sheet2.cell('D22').value = nil
      @sheet2.cell('F22').value = nil    
    }.not_to raise_error
  end
  it 'can be inserted before existing cell and this one is shifted right' do
    @sheet1.B2 = 'test'
    @sheet1.B2.should == 'test'
    inscell = @sheet1.insert_cell_before(2,2)
    inscell.value = 'new'
    @sheet1.B2.should == 'new'
    @sheet1.C2.should == 'test'
    inscell = @sheet1.insert_cell_before(2,4) # should not move cells with data
    @sheet1.C2.should == 'test' 
  end
  it 'Does not ignore rows repeated on every page = header rows (issue 43)' do
    sheet = Rspreadsheet.new('testfile-issue-42-42.ods').sheet(1)
    sheet.A1.should == 'Schedule'
    sheet.B2.should == 'Course'
    sheet.A3.should == 'Teacher'
  end
  it 'Does not ignore cells covered by other merged cells (issue 42)' do
    sheet = Rspreadsheet.new('testfile-issue-42-42.ods').sheet(1)
    sheet.C4.should == 'week2'
    sheet.C5.should == 'week3'
    sheet.C6.should == 'week4'
    sheet.C9.should == 'week7'        
  end
end
