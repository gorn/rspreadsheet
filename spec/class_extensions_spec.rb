require 'spec_helper'
using ClassExtensions


if RUBY_VERSION > '2.1'
  # testing ClassExtensionsForSpec
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

      @m3 = LibXML::XML::Node.new('a')
      @m3 << LibXML::XML::Node.new('i','italic')
       c = LibXML::XML::Node.new('p','paragraph')
       c << LibXML::XML::Node.new('b','boldtext-another')
      @m3 << c
      @m3 << LibXML::XML::Node.new_text('textnode-other')
    end
    it 'can compare nodes' do
      @n.should be === @m
      @n.should_not === @m2
      @n.should be === @m
      @n.should_not === @m2
    end
    it 'has correct text' do
      @n.first_diff(@m).should == nil
      @n.first_diff(nil).inspect.should_not == nil
      @n.first_diff(@m3).inspect.should include 'boldtext-another'
    end
  end

  # testing ClassExtensions
  begin
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
  end
end
