require 'spec_helper'
 
describe Rspreadsheet::Tools do
  it 'Converts correctly cell adresses to coordinates and back' do
    Rspreadsheet::Tools.convert_cell_address_to_coordinates('A1').should == [1,1]
    Rspreadsheet::Tools.convert_cell_address_to_coordinates('C17').should == [17,3]
    Rspreadsheet::Tools.convert_cell_address_to_coordinates('AM1048576').should == [1048576,39]
    Rspreadsheet::Tools.convert_cell_address_to_coordinates('ADA2').should == [2,781]
    Rspreadsheet::Tools.convert_cell_address_to_coordinates('ZZ1').should == [1,702]
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([1,1]).should == 'A1'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([17,3]).should == 'C17'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([1,27]).should == 'AA1'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([1,39]).should == 'AM1'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([1,53]).should == 'BA1'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([1,702]).should == 'ZZ1'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([2,703]).should == 'AAA2'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([1048576,39]).should == 'AM1048576'
    Rspreadsheet::Tools.convert_cell_coordinates_to_address([2,781]).should == 'ADA2' 
    Rspreadsheet::Tools.c2a([2,781]).should == 'ADA2' 
    Rspreadsheet::Tools.c2a(2,781).should == 'ADA2' 
    Rspreadsheet::Tools.a2c('ADA2').should == [2,781]
    Rspreadsheet::Tools.a2c('ADA','2').should == [2,781]
    (1..200).each { |i| Rspreadsheet::Tools.a2c(Rspreadsheet::Tools.c2a(1,i*10)).should == [1,i*10] }
  end
end
