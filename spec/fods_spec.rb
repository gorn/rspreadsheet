require 'spec_helper'
using ClassExtensions if RUBY_VERSION > '2.1'

describe 'Rspreadsheet flat ODS format' do
  before do
    delete_tmpfile(@tmp_filename_fods = '/tmp/testfile.fods')   # delete temp file before tests
    delete_tmpfile(@tmp_filename_ods = '/tmp/testfile.ods')
  end
  after do
    delete_tmpfile(@tmp_filename_fods)
    delete_tmpfile(@tmp_filename_ods)
  end

  it 'can open flat ods testfile and reads its content correctly' do
    book = Rspreadsheet.open($test_filename_fods, format: :fods )
    s = book.worksheets(1)
    (1..10).each do |i|
      s[i,1].should === i
    end
    s[1,2].should === 'text'
    s[2,2].should === Date.new(2014,1,1)
    
    cell = s.cell(6,3)
    cell.format.bold.should == true
    cell = s.cell(6,4)
    cell.format.bold.should == false
    cell.format.italic.should == true
    cell = s.cell(6,5)
    cell.format.italic.should == false
    cell.format.color.should == '#ff3333'
    cell = s.cell(6,6)
    cell.format.color.should_not == '#ff3333'
    cell.format.background_color.should == '#6666ff'
    cell = s.cell(6,7)
    cell.format.font_size.should == '7pt'
  end
  
  it 'does not change when opened and saved again' do
    book = Rspreadsheet.new($test_filename_fods, format: :flat)    # open test file
    book.save(@tmp_filename_fods)                                  # and save it as temp file
    Rspreadsheet::Tools.xml_file_diff($test_filename_fods, @tmp_filename_fods).should be_nil
  end
  
  it 'can be converted to normal format with convert_format_to_normal', :pending do
    book = Rspreadsheet.open($test_filename_fods, format: :flat)
    book.convert_format_to_normal
    book.save_as(@tmp_filename_ods)
    Rspreadsheet::Tools.content_xml_diff($test_filename_fods, @tmp_filename_ods).should be_nil
  end

  it 'pick format automaticaaly' do
    book = Rspreadsheet.open($test_filename_fods)
    book.flat_format?.should == true
    book.save_as(@tmp_filename_fods)
    expect {book = Rspreadsheet.open(@tmp_filename_fods)}.not_to raise_error
    book.normal_format?.should == false
    
    book = Rspreadsheet.open($test_filename_ods)
    book.normal_format?.should == true
    book.save_as(@tmp_filename_ods)
    expect {book = Rspreadsheet.open(@tmp_filename_ods)}.not_to raise_error
    book.flat_format?.should == false
  end

  private
  def delete_tmpfile(afile)
    File.delete(afile) if File.exist?(afile)
  end
  
end
