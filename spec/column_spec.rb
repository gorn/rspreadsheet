require 'spec_helper'
 
describe Rspreadsheet::Cell do
  before do 
    book1 = Rspreadsheet.new
    @sheet = book1.create_worksheet
  end
  it 'contains the right cells' do
    @sheet[1,1] = '11'
    @sheet[1,2] = '12'
    @sheet[2,1] = '21'
    @sheet[2,2] = '22'
    @sheet.column(1).cell(2).value.should eq('21')
    @sheet.column(2).cell(2).value.should eq('22')
    @sheet.column(2).cell(1).should eq(@sheet.cell(1,2))
  end
  
end