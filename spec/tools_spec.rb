require 'spec_helper'
 
describe Rspreadsheet::Tools do
  before do
    @tools = Rspreadsheet::Tools
  end
  it 'converts correctly cell adresses to coordinates' do
    @tools.convert_cell_address_to_coordinates('A1').should == [1,1]
    @tools.convert_cell_address_to_coordinates('C17').should == [17,3]
    @tools.convert_cell_address_to_coordinates('AM1048576').should == [1048576,39]
    @tools.convert_cell_address_to_coordinates('Am1048576').should == [1048576,39]
    @tools.convert_cell_address_to_coordinates('aDa2').should == [2,781]
    @tools.convert_cell_address_to_coordinates('Zz1').should == [1,702]
    @tools.a2c('AdA2').should == [2,781]
    @tools.a2c('ADA','2').should == [2,781]
    @tools.a2c('G','11').should == [11,7]
    @tools.a2c(11, 'G').should == [11,7]
  end
  it 'converts correctly cell coordinates to adresses' do
    @tools.convert_cell_coordinates_to_address([1,1]).should == 'A1'
    @tools.convert_cell_coordinates_to_address([17,3]).should == 'C17'
    @tools.convert_cell_coordinates_to_address([1,27]).should == 'AA1'
    @tools.convert_cell_coordinates_to_address([1,39]).should == 'AM1'
    @tools.convert_cell_coordinates_to_address([1,53]).should == 'BA1'
    @tools.convert_cell_coordinates_to_address([1,702]).should == 'ZZ1'
    @tools.convert_cell_coordinates_to_address([2,703]).should == 'AAA2'
    @tools.convert_cell_coordinates_to_address([1048576,39]).should == 'AM1048576'
    @tools.convert_cell_coordinates_to_address([2,781]).should == 'ADA2' 
    @tools.c2a([2,781]).should == 'ADA2' 
    @tools.c2a(2,781).should == 'ADA2' 
  end
  it 'conversions c2a and a2c are inverse of each other' do
    (1..200).each { |i| @tools.a2c(@tools.c2a(1,i*10)).should == [1,i*10] }
  end
  it 'raises exception when given rubbisch' do
    expect{ @tools.a2c('A1A') }.to raise_error
    expect{ @tools.a2c('1A11') }.to raise_error
    expect{ @tools.a2c('1A11') }.to raise_error
    expect{ @tools.a2c('F','G') }.to raise_error
    expect{ @tools.a2c(5,'G1') }.to raise_error
    expect{ @tools.a2c('G1',5) }.to raise_error
  end
  it 'converts correctly cell adresses given by components to coordinates' do
    @tools.a2c('A','1').should eq [1,1]
    @tools.a2c('C','17').should eq [17,3]
    @tools.a2c('17','C',).should eq [17,3]
    @tools.a2c('17','ZZ',).should eq [17,702]
  end
  it 'given two numbers converts them correctly even when "hidden" in strings' do
    @tools.a2c('3','17').should eq [3,17]
    @tools.a2c(21,'11.0').should eq [21,11]
    @tools.a2c('23',22/2).should eq [23,11]
  end
  it 'can remove attributes from nodes' do
    node = LibXML::XML::Node.new('a')
    @tools.set_ns_attribute(node,'table','ref','123')
    @tools.get_ns_attribute_value(node,'table','ref').should == '123'
    @tools.remove_ns_attribute(node,'table','ref')
    @tools.get_ns_attribute_value(node,'table','ref').should == nil
    @tools.get_ns_attribute_value(node,'table','ref','nic').should == 'nic'    
  end
end
