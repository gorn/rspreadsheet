require 'spec_helper'
using ClassExtensions

describe Array do
  it 'can sum simple array' do
    a = [1,2,3,4]
    a.sum.should == 10
  end
  it 'ignores text and nils while summing' do
    a = [1,nil, nil,2,3,'foo',5.0]
    a.sum.should == 11
    [nil, 'nic'].sum.should == 0
    [].sum.should == 0
  end
end

describe LibXML::XML::Node do
  before do 
    @n = LibXML::XML::Node.new('a')
    @n << LibXML::XML::Node.new('i','italic')
    b = LibXML::XML::Node.new('p','paragraph')
    b << LibXML::XML::Node.new('b','boldtext')
    @n << b
    @n << LibXML::XML::Node.new_text('textnode')
    
    @m = LibXML::XML::Node.new('a')
    @m << LibXML::XML::Node.new('i','italic')
    c = LibXML::XML::Node.new('p','paragraph')
    c << LibXML::XML::Node.new('b','boldtext')
    @m << c
    @m << LibXML::XML::Node.new_text('textnode')
    
    @m2 = LibXML::XML::Node.new('a')
  end
  it 'can compare nodes' do
    @n.should == @m
    @n.should_not == @m2
  end
  it 'has correct elements' do
#     raise @n.first_diff(@m).inspect
  end
end
