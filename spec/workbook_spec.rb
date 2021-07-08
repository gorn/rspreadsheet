require 'spec_helper'

describe Rspreadsheet::Workbook do
  it 'has correct number of sheets' do
    book = Rspreadsheet::Workbook.new($test_filename)
    book.worksheets_count.should == 1
    book.size.should == 1
    book.worksheets(0).should be_nil
    book.worksheets(1).should be_kind_of(Rspreadsheet::Worksheet)
    book.worksheets(2).should be_nil
    book.worksheets(nil).should_not be_nil
    book.worksheet(nil).should be_nil
  end
  it 'freshly created has correctly namespaced xmlnode' do
    @xmlnode = Rspreadsheet::Workbook.new.xmlnode
    @xmlnode.namespaces.to_a.size.should >5
    @xmlnode.namespaces.find_by_prefix('office').should_not be_nil
    @xmlnode.namespaces.find_by_prefix('table').should_not be_nil
    @xmlnode.namespaces.namespace.prefix.should == 'office'
  end
  it 'can create worksheets, and count them' do
    book = Rspreadsheet::Workbook.new()
    book.worksheets_count.should == 0
    book.create_worksheet
    book.worksheets_count.should == 1
    book.create_worksheet
    book.create_worksheet
    book.worksheets_count.should == 3
  end
  it 'nonemptycells behave correctly' do
    book = Rspreadsheet::Workbook.new()
    book.create_worksheet
    @sheet = book.worksheets(1)
    @sheet.cells(3,3).value = 'data'
    @sheet.cells(5,7).value = 'data'
    @sheet.nonemptycells.collect{|c| c.coordinates}.should =~  [[3,3],[5,7]]
  end
  it 'can create named sheets and they are the same as "numbered" ones' do
    book = Rspreadsheet::Workbook.new
    book.create_worksheet('test')
    book.worksheets('test').should eq book.worksheets(1)
    book.create_worksheet('another')
    book.worksheets('another').should eq book.worksheets(2)
  end
  it 'can access sheets with brief syntax' do
    book = Rspreadsheet::Workbook.new
    book.create_worksheet('test')
    book.worksheets('test').should be book.worksheets(1)
    book['test'].should be book.worksheets(1)
    book[1].should be book.worksheets(1)
  end
  it 'can access sheet with alternative syntax and always returns the same object' do
    book = Rspreadsheet::Workbook.new
    book.create_worksheet('test')
    book.create_worksheet('test2')
    sheet = book.worksheets(1)
    book.worksheet(1).should == sheet
    book.sheet(1).should == sheet
    book.sheets(2).should_not == sheet
  end
  it 'can access sheet using negative indexes and returns the same object' do
    book = Rspreadsheet::Workbook.new
    book.create_worksheet('test')
    book.create_worksheet('test2')
    sheet1 = book.worksheets(1)
    sheet2 = book.worksheets(2)
    book.worksheet(-1).should == sheet2
    book.sheet(-2).should == sheet1
    book[-2].should == sheet1
  end
  it 'raises error when attemting to use nonsence index' do
    book = Rspreadsheet::Workbook.new
    expect { book.worksheet(Array.new()) }.to raise_error
  end
end













