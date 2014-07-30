require 'spec_helper'

describe Rspreadsheet::Worksheet do
  before do 
    @sheet = Rspreadsheet.new($test_filename).worksheets[1]
  end
  it 'contains nonempty xml in rows for testfile' do
    @sheet.rows(1).xmlnode.elements.size.should be >1
  end
  it 'freshly created has correctly namespaced xmlnode' do
    @spreadsheet = Rspreadsheet.new
    @spreadsheet.create_worksheet
    @xmlnode = @spreadsheet.worksheets[1].xmlnode
    @xmlnode.namespaces.to_a.size.should >5
    @xmlnode.namespaces.find_by_prefix('office').should_not be_nil
    @xmlnode.namespaces.find_by_prefix('table').should_not be_nil
    @xmlnode.namespaces.namespace.should_not be_nil
    @xmlnode.namespaces.namespace.prefix.should == 'table'
    binding.pry
#     @xmlnode.attributes['name'].namespaces.namespace
  end
 
end
