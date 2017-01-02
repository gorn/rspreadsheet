require 'spec_helper'
 
describe Rspreadsheet::Image do
  before do
    @testbook_images_filename = './spec/testfile2-images.ods'
    @testimage_basename = 'test-image-blue.png'
    @testimage_filename = "./spec/#{@testimage_basename}"
    @sheet = Rspreadsheet.new(@testbook_images_filename).worksheets(1)
  end
  it 'is accesible when included in spreadsheet ', :pending do
    @sheet.images.count.should == 1
    @image = @sheet.images(1)
    @image.name.should == 'Obr-a'
    @sheet.insert_image(@testimage_filename)  ## should it be named this way?
    
    @sheet.images.count.should == 2
    @image = @sheet.images(2)
    @image.filename.should == @testimage_basename ## should it be named this way? - investigate File object
    @image.file.should_be File
    @image.name = 'name1'
    @image.name.should == 'name1'
    @image.name = 'name2'
    @image.name.should != 'name1'
    
    @image.width = '30mm'
    @image.width.should == '30mm'
    @image.height = '31mm'
    @image.height.should == '31mm'
    @image.x = '32mm'
    @image.x.should == '32mm'
    @image.y = '34mm'
    @image.y.should == '34mm'

  end
end