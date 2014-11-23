require 'spec_helper'

describe Rspreadsheet do
  it 'can open ods testfile and reads its content correctly' do
    book = Rspreadsheet.new($test_filename)
    s = book.worksheets(1)
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
    @sheet1 = book1.worksheets(1)
    @sheet2 = book2.worksheets(1)
    
    @sheet1.nonemptycells.each do |cell|       # and test identity
      @sheet2[cell.rowi,cell.coli].should == cell.value
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
    
    @content_xml2.root.first_diff(@content_xml1.root).should be_nil
    @content_xml1.root.first_diff(@content_xml2.root).should be_nil
    
    @content_xml1.root.equals?(@content_xml2.root).should == true
  end
  it 'when open and save file modified, than the file is different' do
    tmp_filename = '/tmp/testfile1.ods'        # first delete temp file
    File.delete(tmp_filename) if File.exists?(tmp_filename)
    book = Rspreadsheet.new($test_filename)    # than open test file
    book.worksheets(1).rows(1).cells(1).value.should_not == 'xyzxyz'
    book.worksheets(1).rows(1).cells(1).value ='xyzxyz'
    book.worksheets(1).rows(1).cells(1).value.should == 'xyzxyz'
    
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
  it 'examples from README file are working' do
    book = Rspreadsheet.open($test_filename)
    sheet = book.worksheets(1)
    sheet.B5 = 'cell value'
    
    sheet.B5.should eq 'cell value'
    sheet[5,2].should eq 'cell value'
    sheet.rows(5).cells(2).value.should eq 'cell value'
    
    expect {
      sheet.F5 = 'text'
      sheet[5,2] = 7
      sheet.cells(5,2).value = 1.78
      
      sheet.cells(5,2).format.bold = true
      sheet.cells(5,2).format.background_color = '#FF0000'
    }.not_to raise_error
    
    sheet.rows(4).cellvalues.sum{|val| val.to_f}.should eq 4+7*4
    sheet.rows(4).cells.sum{ |cell| cell.value.to_f }.should eq 4+7*4

    total = 0
    sheet.rows.each do |row|
      expect {"Sponsor #{row[1]} with email #{row[2]} has donated #{row[3]} USD." }.not_to raise_error
      total += row[1].to_f
    end
    total.should eq 55
  end
end




























