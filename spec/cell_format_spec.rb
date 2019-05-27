require 'spec_helper'

# tests for graphical formats (value formats are elsewhere)
describe Rspreadsheet::Cell do
  before do 
    book1 = Rspreadsheet.new
    @sheet1 = book1.create_worksheet
    book2 = Rspreadsheet.new($test_filename)
    @sheet2 = book2.worksheets(1)
  end
  it 'can have different formats' do
    @cell = @sheet2.cell(6,3)
    @cell.format.bold.should == true
    @cell = @sheet2.cell(6,4)
    @cell.format.bold.should == false
    @cell.format.italic.should == true
    @cell = @sheet2.cell(6,5)
    @cell.format.italic.should == false
    @cell.format.color.should == '#ff3333'
    @cell = @sheet2.cell(6,6)
    @cell.format.color.should_not == '#ff3333'
    @cell.format.background_color.should == '#6666ff'
    @cell = @sheet2.cell(6,7)
    @cell.format.font_size.should == '7pt'
        
    # after fresh create
    @cell.xmlnode.attributes['style-name'].should_not be_nil
  end
  it 'can set formats of the cell in new file' do
    @cell = @sheet1.cell(1,1)
    @cell.value = '1'
    # bold
    @cell.format.bold.should be_falsey
    @cell.format.bold = true
    @cell.format.bold.should be_truthy
    # italic
    @cell.format.italic.should be_falsey
    @cell.format.italic = true
    @cell.format.italic.should be_truthy
    # color
    @cell.format.color.should be_nil
    @cell.format.color = '#AABBCC'
    @cell.format.color.should eq '#AABBCC'
    # background_color
    @cell.format.background_color.should be_nil
    @cell.format.background_color = '#AABBCC'
    @cell.format.style_name.should_not eq 'cell'
    @cell.format.background_color.should eq '#AABBCC'
    # font_size
    @cell.format.font_size.should be_nil
    @cell.format.font_size = '11pt'
    @cell.format.font_size.should eq '11pt'
  end
  it 'setting format in new file detaches the cell' do
    @cell = @sheet1.cell(1,1)
    # bold
    @cell.format.bold.should be_falsey
    @cell.format.bold = true
    @cell.format.bold.should be_truthy
    @cell.mode.should eq :regular

    @cell = @sheet1.cell(2,2)
    @cell.format.background_color = '#ffeeaa'
    @cell.format.background_color.should == '#ffeeaa'
    @cell.mode.should eq :regular
  end
  it 'is possible to manipulate borders of cells' do
    @cell = @sheet1.cell(1,1)
    
    [@cell.format.top,@cell.format.left,@cell.format.right,@cell.format.bottom].each do |border|
      border.style = 'dashed'
      border.style.should == 'dashed'
      border.width = 0.5 
      border.width.should == 0.5
      border.color = '#005500'
      border.color.should == '#005500'
    end    
  end
  it 'returns correct border parameters for the cell' do
    @sheet2.cell('C8').format.top.style.should == 'solid'
    @sheet2.cell('E8').format.left.color.should == '#ff3333'
    @sheet2.cell('E8').format.left.style.should == 'solid'
    @sheet2.cell('F8').format.top.color.should == '#009900'
    @sheet2.cell('F8').format.top.style.should == 'dotted'
  end
  it 'modifies borders correctly' do
    ## initially solid everywhere
    @sheet2.cell('C8').format.top.style.should == 'solid'
    @sheet2.cell('C8').format.bottom.style.should == 'solid'
    @sheet2.cell('C8').format.left.style.should == 'solid'
    @sheet2.cell('C8').format.right.style.should == 'solid'
    ## change top and right to dotted and observe
    @sheet2.cell('C8').format.top.style = 'dotted'
    @sheet2.cell('C8').format.right.style = 'dotted'
    @sheet2.cell('C8').format.bottom.style.should == 'solid'
    @sheet2.cell('C8').format.left.style.should == 'solid'
    @sheet2.cell('C8').format.top.style.should == 'dotted'
    @sheet2.cell('C8').format.right.style.should == 'dotted'
  end
  it 'deletes borders correctly', :pending=> 'consider how to deal with deleted borders' do
    @cell = @sheet1.cell(1,1)
    
    [@cell.format.top,@cell.format.left,@cell.format.right,@cell.format.bottom].each do |border|
      border.style = 'dashed'
      border.should_not be_nil
      border.delete
      border.should be_nil
    end 
    
    # delete right border in existing file and observe
    @sheet2.cell('C8').format.right.delete
    @sheet2.cell('C8').format.right.should == nil
  end
  
  it 'can delete borders in many ways', :pending => 'consider what syntax to support' do
    @cell=@sheet2.cell('C8')
    @cell.border_right.should_not be_nil
    @cell.border_right.delete
    @cell.border_right.should be_nil
    
    @cell.border_left.should_not be_nil
    @cell.border_left = nil
    @cell.border_left.should be_nil
    
    @cell.format.top.should_not_be_nil
    @cell.format.top.style = 'none'
    @cell.border_top.should_not be_nil ## ?????
  end
  
  it 'does not have issue 40 - coloring other cells, when they have similar borders' do
    sheet = Rspreadsheet.open('./spec/testfile-issue-40.ods').worksheet(1)
    sheet[1,1] = 'ahoj'
    sheet.cell(1,1).format.background_color = '#FF0000'

    sheet.cell(1,1).format.background_color.should == '#FF0000'
    sheet.cell('B3').format.background_color.should_not == '#FF0000'
  end
  
  it 'automatically creates new style, if a style is automatic, some of its attributes changes and there are several cells pointing to it', :focus do
    sheet = Rspreadsheet.open('./spec/testfile-issue-40.ods').worksheet(1)
    cell1 = sheet.cell(1,1)
    cell2 = sheet.cell('B3')
    
    cell1.xmlnode.attributes['style-name'].should == cell2.xmlnode.attributes['style-name']
    cell1.format.background_color = '#FF0000'
    
    cell1.xmlnode.attributes['style-name'].should_not == cell2.xmlnode.attributes['style-name']
  end
  
  it 'recognizes correctly that style is shared with other cell' do
    sheet = Rspreadsheet.open('./spec/testfile-issue-40.ods').worksheet(1)
    
    sheet.cell(1,1).format.style_shared_count.should == 2
    sheet.cell(1,1).format.style_shared?.should == true
  end
  
end
