require 'spec_helper'
using ClassExtensions if RUBY_VERSION > '2.1'

describe 'Rspreadsheet flat ODS format' do
  before do
    @tmp_filename = '/tmp/testfile.ods'        # delete temp file before tests
    delete_tmpfile(@tmp_filename)
  end
  after do
    delete_tmpfile(@tmp_filename) # delete temp file after tests
  end

  it 'can open flat ods testfile and reads its content correctly' do
    book = Rspreadsheet.open($test_filename_fods, format: :fods )
    s = book.worksheets(1)
    (1..10).each do |i|
      s[i,1].should === i
    end
    s[1,2].should === 'text'
    s[2,2].should === Date.new(2014,1,1)
  end
  
  it 'can open and save flast ods file, and saved file is exactly same as original' do
    book = Rspreadsheet.new($test_filename_fods)    # open test file
    book.save(@tmp_filename)                    # and save it as temp file
    
    # now compare them
    @content_xml1 = Zip::File.open($test_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_xml2 = Zip::File.open(@tmp_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    
    @content_xml2.root.first_diff(@content_xml1.root).should be_nil
    @content_xml1.root.first_diff(@content_xml2.root).should be_nil
    @content_xml1.root.to_s.should == @content_xml2.root.to_s
  end

  private
  def delete_tmpfile(afile)
    File.delete(afile) if File.exist?(afile)
  end
  
end
