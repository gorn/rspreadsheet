require 'spec_helper'

describe Rspreadsheet::XmlTiedArray do
  before do
#     construct @xmlnode with value 
#        <table:table-row>
#          <table:table-cell table:number-columns-repeated="5"/>
#          <table:table-cell table:number-columns-repeated="3"/>
#          <table:table-cell office:value-type="string">
#            <text:p>text content</text:p>
#          </table:table-cell>
#        </table:table-row>
    @xmlnode = Rspreadsheet::Tools.create_ns_node('table-row','table')
    xmlsubnode = Rspreadsheet::Tools.create_ns_node('table-cell','table')
    Rspreadsheet::Tools.set_ns_attribute(xmlsubnode,'table','number-columns-repeated','5')
    @xmlnode << xmlsubnode
    xmlsubnode = Rspreadsheet::Tools.create_ns_node('table-cell','table')
    Rspreadsheet::Tools.set_ns_attribute(xmlsubnode,'table','number-columns-repeated','3')
    @xmlnode << xmlsubnode
    xmlsubnode = Rspreadsheet::Tools.create_ns_node('table-cell','table')
    Rspreadsheet::Tools.set_ns_attribute(xmlsubnode,'office','value-type','string')
    xmlsub2node = Rspreadsheet::Tools.create_ns_node('p','text')
    xmlsub2node << 'text content'
    xmlsubnode << xmlsub2node 
    @xmlnode << xmlsubnode
    @xarr = Rspreadsheet::XmlTiedArray.new(@xmlnode, xml_repeated_attribute: 'number-columns-repeated', xml_items_node_name: 'table-cell' )
  end
  it 'can get items out correctly' do
    @item = @xarr.get_item(2)
    @item.repeated?.should be true
    @item.index.should == 2
#     @xarr = XmlTiedArray.new(nil,@xmlnode)
  end
 
end