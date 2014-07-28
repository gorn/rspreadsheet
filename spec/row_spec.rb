require 'spec_helper'

describe Rspreadsheet::Row do
  before do 
    book1 = Rspreadsheet.new
    @sheet1 = book1.create_worksheet
  end
  it 'allows access to cells in a row' do
    @row = Rspreadsheet::SingleRow.new(nil,1)
    @row.cells(3).value = 3
    @row.cells(3).value.should == 3
  end
  it 'cells in row are settable through sheet' do
    (2..5).each { |i| @sheet1[7,i] = i }
    (2..5).each { |i| 
      a = @sheet1.rows(7)
      c = a.cells(i)
      c.value.should == i 
    } 
  end
  it 'is found even in empty sheets' do
    @sheet1.rows(5).should be_kind_of(Rspreadsheet::Row)
    @sheet1.rows(25).should be_kind_of(Rspreadsheet::Row)
    @sheet1.cells(10,10).value = 'ahoj'
    @sheet1.rows(9).should be_kind_of(Rspreadsheet::Row)
    @sheet1.rows(10).should be_kind_of(Rspreadsheet::Row)
    @sheet1.rows(11).should be_kind_of(Rspreadsheet::Row)
  end
  it 'detaches automatically row and creates new repeated rows when needed' do
    @sheet1.rows(5).detach
    @sheet1.rows(5).repeated?.should == false
    @sheet1.rows(5).repeated.should == 1
    @sheet1.rows(5).xmlnode.should_not == nil
    
    @sheet1.rows(3).repeated?.should == true
    @sheet1.rows(3).repeated.should == 4
    @sheet1.rows(3).xmlnode.should_not == nil
    @sheet1.rows(3).xmlnode['table:number-rows-repeated'].should == '4'
    
    @sheet1.rows(5).style_name = 'newstylename'
    @sheet1.rows(5).xmlnode.attributes['table:style-name'].should == 'newstylename'

    @sheet1.rows(17).style_name = 'newstylename2'
    @sheet1.rows(17).xmlnode.attributes['table:style-name'].should == 'newstylename2'
  end
  it 'ignores negative any zero row indexes' do
    @sheet1.rows(0).should be_nil
    @sheet1.rows(-78).should be_nil
  end
  it 'has correct rowindex' do
    @sheet1.rows(5).detach
    (4..6).each do |i|
      @sheet1.rows(i).row.should == i
    end
  end
  it 'can open ods testfile and read its content' do
    book = Rspreadsheet.new($test_filename)
    s = book.worksheets[1]
    (1..10).each do |i|
      s.row(i).should be_kind_of(Rspreadsheet::Row)
      s.row(i).repeated.should == 1
      s.row(i).used_range.size>0
    end
    s[1,2].should === 'text'
    s[2,2].should === Date.new(2014,1,1)
  end
  it 'normalizes to itself if single line' do
    @sheet1.rows(5).detach
    @sheet1.rows(5).cell(4).value='test'
    @sheet1.rows(5).normalize
    @sheet1.rows(5).cell(4).value.should == 'test'
  end
end

 
