require 'spec_helper'
using ClassExtensions if RUBY_VERSION > '2.1'

describe Rspreadsheet do
  before do
    @tmp_filename = '/tmp/testfile.ods'
    File.delete(@tmp_filename) if File.exists?(@tmp_filename)  # delete temp file
  end
  
  
  it 'when saved to file is identical' do
    spreadsheet = Rspreadsheet.new($test_filename)                 # open a file
    spreadsheet.save(@tmp_filename)                                # and save spreadsheet as temp file
    
    # now compare content saved file to original
    contents_of_files_should_be_identical($test_filename,@tmp_filename).should == true
  end
  
  it 'when saved to file via save_to_io is identical' do
    @tmp_filename = '/tmp/testfile4.ods'
    spreadsheet = Rspreadsheet.new($test_filename)                 # open a file
    File.open(@tmp_filename, 'w+') do |file|
       spreadsheet.save_to_io(file)                                   # and save spreadsheet as temp file
    end
    
    # now compare content saved file to original
    contents_of_files_should_be_identical($test_filename,@tmp_filename)
  end
  
  it 'can be saved to IO object' do
    @tmp_filename = '/tmp/testfile5.ods'
    spreadsheet = Rspreadsheet.new($test_filename_images)                 # open a file
    
    stringio = StringIO.new
    spreadsheet.save_to_io(stringio)
#     stringio.size.should > 300000
    
    # save it to temp file
    File.open(@tmp_filename, "w") do |f|
      f.write stringio.read                                  
    end
     
    contents_of_files_should_be_identical($test_filename_images,@tmp_filename)
  end
  

end

def xml_from_entry(zip,entryname)
  LibXML::XML::Document.io(zip.get_input_stream(entryname))
end

def contents_of_files_should_be_identical(filename1,filename2)
  Zip::File.open(filename1) do |zip|
    @content1_xml = xml_from_entry(zip,'content.xml')
    @manifest1_xml = xml_from_entry(zip,'META-INF/manifest.xml')
    @images1_count = zip.glob('Pictures/**').count
  end
  
  Zip::File.open(filename2) do |zip|
    @content2_xml = xml_from_entry(zip,'content.xml')
    @manifest2_xml = xml_from_entry(zip,'META-INF/manifest.xml')
    @images2_count = zip.glob('Pictures/**').count
  end
  
  @images1_count.should == @images2_count
  xmls_should_be_identical(@content1_xml,@content2_xml) 
  xmls_should_be_identical(@manifest1_xml,@manifest2_xml)
end

def xmls_should_be_identical(xml1,xml2) 
  xml2.root.first_diff(xml1.root).should be_nil
  xml1.root.first_diff(xml2.root).should be_nil
  xml1.root.to_s.should == xml2.root.to_s
end
