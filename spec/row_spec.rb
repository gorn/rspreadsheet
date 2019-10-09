require 'spec_helper'

describe Rspreadsheet::Row do
  before do
    @sheet1 = Rspreadsheet.new.create_worksheet
    @book2 = Rspreadsheet.new($test_filename)
    @sheet2 = @book2.worksheets(1)
  end
  it 'allows access to cells in a row' do
    @row = @sheet2.rows(1)
    @c = @row.cells(1)
    @c.value = 3
    @c.value.should == 3
  end
  it 'allows access to cells using different syntax' do
    @row = @sheet1.rows(1)
    @row.cells(1).value = 77
    @row.cells(1).value.should == 77
    @row.cells.first.value.should == 77
    @row.cells.first.value = 88
    @row.cells(1).value.should == 88
    @row[1]= 99
    @row.cells(1).value.should == 99
  end
  it 'can be detached and changes to unrepeated if done' do
    @row = @sheet1.rows(5)
    @row.xmlnode.andand.name.should_not == 'table-row'
    @row2 = @row.detach
    @row2.xmlnode.name.should == 'table-row'
    @row2.is_repeated?.should == false
  end
  it 'is the synchronized object, now matter how you access it' do
    @row1 = @sheet1.rows(5)
    @row2 = @sheet1.rows(5)
    @row1.should equal(@row2)

    @sheet1.rows(5).cells(2).value = 'nejakydata'
    @row1 = @sheet1.rows(5)
    @row2 = @sheet1.rows(5)
    @row1.should equal(@row2)

  end
  it 'cells in row are settable through sheet' do
    @sheet1.rows(9).cells(1).value = 2
    @sheet1.rows(9).cells(1).value.should == 2

    @sheet1.rows(7).cells(1).value = 7
    @sheet1.rows(7).cells(1).value.should == 7

    @sheet1.rows(5).cells(1).value = 5
    @sheet1.rows(5).cells(1).value.should == 5

    (2..5).each { |i| @sheet1.rows(3).cells(i).value = i }
    (2..5).each { |i|
      a = @sheet1.rows(3)
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
  it 'detachment creates correct repeated groups' do
    @sheet1.rows(5).detach
    @sheet1.rows(5).repeated?.should == false
    @sheet1.rows(5).repeated.should == 1
    @sheet1.rows(5).xmlnode.should_not == nil

    @sheet1.rows(3).repeated?.should == true
    @sheet1.rows(3).repeated.should == 4
    @sheet1.rows(3).xmlnode.should_not == nil
    @sheet1.rows(3).xmlnode.attributes.get_attribute_ns("urn:oasis:names:tc:opendocument:xmlns:table:1.0",'number-rows-repeated').value.should == '4'
  end
  it 'detachment assigns correct namespaces to node' do
    @sheet1.rows(5).detach
    @xmlnode = @sheet1.rows(5).xmlnode
    @xmlnode.namespaces.to_a.size.should >5
    @xmlnode.namespaces.namespace.should_not be_nil
    @xmlnode.namespaces.namespace.prefix.should == 'table'
  end
  it 'by assigning value, the repeated row is automatically detached' do
    @sheet1.rows(15).detach

    @sheet1.rows(2).repeated?.should == true
    @sheet1.rows(2).cells(2).value = 'nejakydata'
    @sheet1.rows(2).repeated?.should == false

    @sheet1.rows(22).repeated?.should == true
    @sheet1.rows(22).cells(7).value = 'nejakydata'
    @sheet1.rows(22).repeated?.should == false
  end
  it 'styles can be assigned to rows' do
    @sheet1.rows(5).detach
    table_ns_href = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"

    @sheet1.rows(5).style_name = 'newstylename'
    @sheet1.rows(5).xmlnode.attributes.get_attribute_ns(table_ns_href,'style-name').value.should == 'newstylename'

    @sheet1.rows(17).style_name = 'newstylename2'
    @sheet1.rows(17).xmlnode.attributes.get_attribute_ns(table_ns_href,'style-name').value.should == 'newstylename2'
  end
  it 'returns nil on negative and zero row indexes or raises exception depending on configuration' do
    pom = Rspreadsheet.raise_on_negative_coordinates
    # default is to raise exception
    expect { @sheet1.rows(0) }.to raise_error
    expect { @sheet1.rows(-78) }.to raise_error

    # if configured, it needs to ignore it (and return nil)
    Rspreadsheet.raise_on_negative_coordinates = false
    @sheet1.rows(0).should be_nil
    @sheet1.rows(-78).should be_nil

    Rspreadsheet.raise_on_negative_coordinates = pom  # reset the setting back
  end
  it 'has correct rowindex' do
    @sheet1.rows(5).detach
    (4..6).each do |i|
      @sheet1.rows(i).rowi.should == i
    end
  end
  it 'can open ods testfile and read its content' do
    book = Rspreadsheet.new($test_filename)
    s = book.worksheets(1)
    (1..10).each do |i|
      s.rows(i).should be_kind_of(Rspreadsheet::Row)
      s.rows(i).repeated.should == 1
      s.rows(i).used_range.size>0
    end
    s[1,2].should === 'text'
    s[2,2].should === Date.new(2014,1,1)
  end
  it 'cell manipulation does not contain attributes without namespace nor doubled attributes' do # inspired by a real bug regarding namespaces manipulation
    @sheet2.rows(1).xmlnode.attributes.each { |attr| attr.ns.should_not be_nil}
    @sheet2.rows(1).cells(1).value.should_not == 'xyzxyz'
    @sheet2.rows(1).cells(1).value ='xyzxyz'
    @sheet2.rows(1).cells(1).value.should == 'xyzxyz'

    ## attributes have namespaces
    @sheet2.rows(1).xmlnode.attributes.each { |attr| attr.ns.should_not be_nil}

    ## attributes are not douubled
    xmlattrs = @sheet2.rows(1).xmlnode.attributes.to_a.collect{ |a| [a.ns.andand.prefix.to_s,a.name].reject{|x| x.empty?}.join(':')}
    duplication_hash = xmlattrs.inject(Hash.new(0)){ |h,e| h[e] += 1; h }
    duplication_hash.each { |k,v| v.should_not >1 }  # should not contain duplicates
  end
  it 'out of bound automagically generated row pick up values when created' do
    @row1 = @sheet1.rows(23)
    @row2 = @sheet1.rows(23)
    @row2.cells(5).value = 'hojala'
    @row1.cells(5).value.should == 'hojala'
  end
  it 'nonempty cells work properly' do
    nec = @sheet2.rows(1).nonemptycells
    nec.collect{ |c| c.coordinates}.should == [[1,1],[1,2]]

    nec = @sheet2.rows(19).nonemptycells
    nec.collect{ |c| c.coordinates}.should == [[19,6]]
  end
  it 'is the same object no matter when you access it' do
    @row1 = @sheet2.rows(5)
    @row2 = @sheet2.rows(5)
    @row1.should equal(@row2)
  end
  it 'reports good range of coordinates for repeated rows' do
    @row2 = @sheet2.rows(15)
    @row2.mode.should == :repeated
    @row2.range.should == (14..18)

    @sheet1.rows(15).detach

    @sheet1.rows(2).repeated?.should == true
    @sheet1.rows(2).range.should  == (1..14)

    @sheet1.rows(22).repeated?.should == true
    @sheet1.rows(22).range.should  == (16..Float::INFINITY)
  end
  it 'shifts rows if new one is added in the middle' do
    @sheet1.rows(15).detach
    @sheet1.rows(20).detach
    @row = @sheet1.rows(16)

    @row.range.should == (16..19)
    @row.rowi.should == 16

    @sheet1.add_row_above(7)
    @sheet1.rows(17).range.should == (17..20)
    @row.range.should == (17..20)
    @row.rowi.should == 17
    @sheet1.rows(17).should equal(@row)

    @row.add_row_above
    @row.rowi.should == 18
  end
  it 'inserted has correct class' do # based on real error
    @sheet2.add_row_above(1)
    @sheet2.rows(1).should be_kind_of(Rspreadsheet::Row)
  end
  it 'inserted is empty even is surrounded by nonempty rows' do
    @sheet2.row(4).cells.size.should > 1
    @sheet2.row(5).cells.size.should == 1
    @row5 = @sheet2.row(5)
    @sheet2.add_row_above(5)
    @sheet2.row(4).cells.size.should > 1
    @sheet2.row(5).cells.size.should == 0
    @sheet2.row(6).should == @row5
  end

  it 'can be deleted' do
    @sheet1[15,4]='data'
    @row = @sheet1.rows(15)
    @row.invalid_reference?.should be false
    @row.delete
    @row.invalid_reference?.should be true
    expect { @row.cells(4) }.to raise_error
    @sheet1.rows(15).invalid_reference?.should be false # this is former line 16
  end
  it 'shifts rows if row is deleted' do
    @row = @sheet1.rows(15)
    @row[1] = 'data1'
    @row[1].should eq 'data1'
    @row.rowi.should == 15

    @sheet1.rows(7).delete

    @row[1].should eq 'data1'
    @row.rowi.should == 14

    @sheet1.rows(14).should be @row
    @sheet1.rows(14).cells(1).value.should eq 'data1'
  end
  it 'has no cells when empty' do
    @row = @sheet1.rows(2)
    @row.cells.size.should == 0
  end
  it 'has 1 cells when assigned to first one' do
    @row = @sheet1.rows(2)
    @row.cells(1).value = 'foo'
    @row.cells.size.should == 1
  end
  it 'can be mass assigned by cellvalues= method' do
    @row = @sheet1.rows(2)
#     @row.size.should == 0
    @row.cellvalues= ['January',nil,3]
    @row.size.should == 3
    @row.cells(1).value.should == 'January'
    @row.cells(2).blank?.should be_truthy
    @row.cells(3).value.should == 3
    @row.cellvalues = [1]
    @row.size.should == 1
    @row.cells(1).value.should == 1
    @row.cells(3).blank?.should be_truthy
  end
  it 'can be mass assigned by sheet[]= method' do
    @sheet1[1] = ['foo','baz']
  end
  it 'cells can be wiped out by truncate method' do
    @row = @sheet1.rows(1)
    @row[1]='data'
    @row[3]='data3'
    @row.size.should == 3
    @row.truncate
    @row.size.should == 0
  end
  it 'remembers its parent correctly' do
    @row = @sheet1.rows(5)
    @row.worksheet.should == @sheet1
  end
  it 'does not skip header rows (issue #43)' do
    @sheet = Rspreadsheet.open('./spec/testfile3-header_rows_and_cells.fods').worksheet(1)
    @sheet.A1.should == 'Cell in header row'
    @sheet.A1.should_not == 'This is first nonheader row, but it is in colheader'
    @sheet.B3.should == 'First completely nonheader cell'
  end
  it 'can be cloned to other row' do
    @sheet2.row(5)[1].should == 5
    @sheet2.row(6)[1].should == 6

    @sheet2.row(6)[1].should_not == 4
    @sheet2.row(4).clone_above_row(6)
    @sheet2.row(4)[1].should == 4
    @sheet2.row(5)[1].should == 5
    @sheet2.row(6)[1].should == 4
    @sheet2.row(7)[1].should == 6
    @sheet2.row(6)[1].should == @sheet2.row(4)[1]
    @sheet2.row(6)[2].should == @sheet2.row(4)[2]
    @sheet2.row(6)[3].should == @sheet2.row(4)[3]
    @sheet2.row(6).cell(2).formula.should == @sheet2.row(4).cell(2).formula
  end
end


