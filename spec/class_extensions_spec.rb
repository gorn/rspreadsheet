require 'spec_helper'

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
    end
    it 'can compare nodes' do
      @n.to_s.should be == @m.to_s
      @n.to_s.should_not == @m2.to_s
    end
    it 'has correct elements' do
  #     raise @n.first_diff(@m).inspect
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
