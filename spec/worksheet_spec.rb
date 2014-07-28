require 'spec_helper'

describe Rspreadsheet::Worksheet do
  before do 
    @sheet = Rspreadsheet.new($test_filename).worksheets[1]
  end
  it 'contains nonempty xml in rows for testfile' do
    @sheet.rows(1).xmlnode.elements.size.should be >1
  end

end 
 