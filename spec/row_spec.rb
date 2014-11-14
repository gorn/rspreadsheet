require 'spec_helper'

describe Rspreadsheet::Row do
  before do 
    @sheet1 = Rspreadsheet.new.create_worksheet
    @sheet2 = Rspreadsheet.new($test_filename).worksheets[1]
  end
  it 'allows access to cells in a row' do
    @row = @sheet2.rows(1)
    @c = @row.cells(1)
    @c.value = 3
    @c.value.should == 3 
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
  it 'ignores negative any zero row indexes' do
    @sheet1.rows(0).should be_nil
    @sheet1.rows(-78).should be_nil
  end
  it 'has correct rowindex' do
    @sheet1.rows(5).detach
    (4..6).each do |i|
      @sheet1.rows(i).rowi.should == i
    end
  end
  it 'can open ods testfile and read its content' do
    book = Rspreadsheet.new($test_filename)
    s = book.worksheets[1]
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
    
    @sheet1.insert_row_above(7)
    @sheet1.rows(17).range.should == (17..20)
    @row.range.should == (17..20)
    @sheet1.rows(17).should equal(@row)    
  end
  it 'inserted has correct class' do # based on real error
    @sheet2.insert_row_above(1)
    @sheet2.rows(1).should be_kind_of(Rspreadsheet::Row)
  end
end

 
