require 'spec_helper'
 
describe Rspreadsheet::Image do
  before do
    @testfile_filename = $test_filename_images
    @tmp_testfile_filename = '/tmp/testfile2-image.ods'
    File.delete(@tmp_testfile_filename) if File.exists?(@tmp_testfile_filename) # delete temp file

    @testimage_filename  = './spec/test-image-blue.png'
    @testimage2_filename = './spec/test-image.png'
    @workbook = Rspreadsheet.new(@testfile_filename)
    @sheet  = @workbook.worksheets(1)
    @sheet2 = @workbook.worksheets(2)
  end
  after do
    File.delete(@tmp_testfile_filename) if File.exists?(@tmp_testfile_filename) # delete temp file
  end
  it 'is accesible when included in spreadsheet', :xpending do
    @sheet.images_count.should == 1
    @image = @sheet.images(1)
    
    @image.name.should == 'Obr-a'
    @sheet.insert_image(@testimage_filename)  ## should it be named this way?
    
    @sheet.images.count.should == 2
    @image = @sheet.images(2)
    @image.original_filename.should == @testimage_filename ## should it be named this way? - investigate File object
    @image.name = 'name1'
    @image.name.should == 'name1'
    @image.name = 'name2'
    @image.name.should_not == 'name1'
    
    @image.width = '30mm'
    @image.width.should == '30mm'
    @image.height = '31mm'
    @image.height.should == '31mm'
    @image.x = '32mm'
    @image.x.should == '32mm'
    @image.y = '34mm'
    @image.y.should == '34mm'
  end
  it 'can be inserted in sheet without any pictures' do
    @sheet2.insert_image_to('10mm','10mm',@testimage_filename)
  end
  
  it 'can be inserted to a sheet and moved around' do
    x,y = '15mm', '17mm'
    @image = @sheet.insert_image_to(x,y,@testimage_filename)
    # moving image
    @image.x = '21mm'
    @image.x.should == '21mm'
    @image.move_to('51mm','52mm')
    @image.y.should == '52mm'
    # resizing image
    @image.width = '30mm'
    @image.height = '30mm'
    @image.width.should == '30mm'
    # copying image
    @sheet.images_count.should == 2
    @sheet2.images_count.should == 0
    @image.copy_to(x,y,@sheet2)
    @sheet.images_count.should == 2
    @sheet2.images_count.should == 1
  end
  
  it 'can be inserted into file and is saved correctly to it' do
    tmp_test_image = '/tmp/test-image.png'
    
    # create new file, insert image into it and save it
    book = Rspreadsheet.new
    @sheet = book.add_worksheet
    @sheet.images_count.should == 0
    book.worksheets(1).insert_image_to('10mm','10mm',@testimage2_filename)
    @sheet.images_count.should == 1
    book.save(@tmp_testfile_filename)

    # reopen it and check the contents
    book2 = Rspreadsheet.new(@tmp_testfile_filename)
    @sheet2 = book2.worksheets(1)
    @sheet2.images_count.should == 1
    @image = @sheet2.images(1)
    
    File.delete(tmp_test_image) if File.exists?(tmp_test_image)
    Zip::File.open(@tmp_testfile_filename) do |zip|  ## TODO: this is UGLY - it should not be extracting contents here
      zip.extract(@image.internal_filename,tmp_test_image)
    end
    File.binread(tmp_test_image).unpack("H*").should == File.binread(@testimage2_filename).unpack("H*")
  end
  
  it 'generates internal_filename on save randomly and they are different' do
    i1 = @sheet.insert_image(@testimage_filename)
    i2 = @sheet.insert_image(@testimage2_filename)
    i3 = @sheet.insert_image(@testimage2_filename)
    i1.move_to('40mm','40mm')
    i2.move_to('40mm','40mm')
    i3.move_to('40mm','40mm')
    @workbook.save(@tmp_testfile_filename)
    
    i1.internal_filename.should_not == i2.internal_filename
    i1.internal_filename.should_not == i3.internal_filename
    i2.internal_filename.should_not == i3.internal_filename
  end
  
#   it 'has dimensions defaulting to size of the image once it is inserted' do
#   end
  
end