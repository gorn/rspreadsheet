require 'spec_helper'
using ClassExtensions if RUBY_VERSION > '2.1'

describe Rspreadsheet do
  before do
    @tmp_filename = '/tmp/testfile.ods'        # delete temp file before tests
    File.delete(@tmp_filename) if File.exist?(@tmp_filename)
  end
  after do
    File.delete(@tmp_filename) if File.exist?(@tmp_filename) # delete temp file after tests
  end
  
  it 'can open ods testfile and reads its content correctly' do
    book = Rspreadsheet.new($test_filename)
    s = book.worksheets(1)
    (1..10).each do |i|
      s[i,1].should === i
    end
    s[1,2].should === 'text'
    s[2,2].should === Date.new(2014,1,1)
  end
  it 'can open and save file, and saved file has same cells as original' do
    
    book = Rspreadsheet.new($test_filename)     # open test file
    book.save(@tmp_filename)                    # and save it as temp file
    
    book1 = Rspreadsheet.new($test_filename)   # now open both again
    book2 = Rspreadsheet.new(@tmp_filename)
    @sheet1 = book1.worksheets(1)
    @sheet2 = book2.worksheets(1)
    
    @sheet1.nonemptycells.each do |cell|       # and test identity
      @sheet2[cell.rowi,cell.coli].should == cell.value
    end
  end

  it 'can open and save file, and saved file is exactly same as original' do
    book = Rspreadsheet.new($test_filename)    # open test file
    book.save(@tmp_filename)                    # and save it as temp file
    
    # now compare them
    @content_xml1 = Zip::File.open($test_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_xml2 = Zip::File.open(@tmp_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    
    @content_xml2.root.first_diff(@content_xml1.root).should be_nil
    @content_xml1.root.first_diff(@content_xml2.root).should be_nil
    @content_xml1.root.to_s.should == @content_xml2.root.to_s
  end

  it 'when open and save file modified, than the file is different' do
    book = Rspreadsheet.new($test_filename)    # open test file
    book.worksheets(1).rows(1).cells(1).value.should_not == 'xyzxyz'
    book.worksheets(1).rows(1).cells(1).value ='xyzxyz'
    book.worksheets(1).rows(1).cells(1).value.should == 'xyzxyz'
    
    book.save(@tmp_filename)                    # and save it as temp file
    
    # now compare them
    @content_doc1 = Zip::File.open($test_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_doc2 = Zip::File.open(@tmp_filename) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    @content_doc1.eql?(@content_doc2).should == false
  end
  it 'can create file' do
    book = Rspreadsheet.new
  end
  it 'can create new worksheet' do
    book = Rspreadsheet.new
    book.create_worksheet
  end
  it 'examples from README file are working' do
    Rspreadsheet.open($test_filename).save(@tmp_filename)
    @testimage_filename  = './spec/test-image-blue.png'
    def puts(*par); end # supress puts in the example
    expect do
      book = Rspreadsheet.open(@tmp_filename)
      sheet = book.worksheets(1)

      # get value of a cell B5 (there are more ways to do this)
      sheet.B5                       # => 'cell value'
      sheet[5,2]                     # => 'cell value'
      sheet.row(5).cell(2).value   # => 'cell value'

      # set value of a cell B5
      sheet.F5 = 'text'
      sheet[5,2] = 7
      sheet.cell(5,2).value = 1.78

      # working with cell format
      sheet.cell(5,2).format.bold = true
      sheet.cell(5,2).format.background_color = '#FF0000'

      # calculate sum of cells in row
      sheet.row(5).cellvalues.sum
      sheet.row(5).cells.sum{ |cell| cell.value.to_f }

      # or set formula to a cell
      sheet.cell('A1').formula='=SUM(A2:A9)'

      # insert company logo to the file
      sheet.insert_image_to('10mm','15mm',@testimage_filename)
      
      # iterating over list of people and displaying the data
      total = 0
      sheet.rows.each do |row|
        puts "Sponsor #{row[1]} with email #{row[2]} has donated #{row[3]} USD."
        total += row[3].to_f
      end
      puts "Totally fundraised #{total} USD"

      # saving file
      book.save
      book.save('/tmp/different_filename.ods')
    end.not_to raise_error
    File.delete('/tmp/different_filename.ods') if File.exist?('/tmp/different_filename.ods') # delete after tests
  end
  it 'examples from advanced syntax GUIDE are working' do
    def p(*par); end # supress p in the example
    expect do 
      book = Rspreadsheet::Workbook.new
      sheet = book.create_worksheet 'Top icecreams'

      sheet[1,1] = 'My top 5'
      p sheet[1,1].class # => String
      p sheet[1,1] # => "My top 5"
      
      # These are all the same values - alternative syntax
      p sheet.rows(1).cells(1).value
      p sheet.cells(1,1).value
      p sheet.A1
      p sheet[1,1]
      p sheet['A1']
      p sheet.cells('A1').value
      
      # How to inspect/manipulate the Cell object
      sheet.cells(1,1) # => Rspreadsheet::Cell
      sheet.cells(1,1).format
      sheet.cells(1,1).format.font_size = '15pt'
      sheet.cells(1,1).format.bold = true
      p sheet.cells(1,1).format.bold? # => true
      
      # There are the same assigmenents
      value = 1.234
      sheet.A1 = value
      sheet[1,1]= value
      sheet.cells(1,1).value = value
      
      p sheet.A1.class # => Rspreadsheet::Cell
      
      # relative cells
      sheet.cells(4,7).relative(-1,0) # => cell 3,7
      
      # build the top five list (these features are not implemented yet)
#       (1..5).each { |i| sheet[i,1] = i }
#       sheet.columns(1).format.bold = true
#       sheet.cells[2,1..5] = ['Vanilla', 'Pistacia', 'Chocolate', 'Annanas', 'Strawbery']
#       sheet.columns(1).cells(1).format.color = :red
      
      book.save('/tmp/testfile.ods')
    end.not_to raise_error
  end
  it 'can save file to io stream and the content is the same as when saving to file' do
    book = Rspreadsheet.new($test_filename)    # open test file
    
    File.open(@tmp_filename, 'w') do |file| 
      file.write(book.to_io.read) # this does not work
    end
  
    book1 = Rspreadsheet.new($test_filename)   # now open both again
    book2 = Rspreadsheet.new(@tmp_filename)
    @sheet1 = book1.worksheets(1)
    @sheet2 = book2.worksheets(1)
    
    @sheet1.nonemptycells.each do |cell|       # and test if they are identical
      @sheet2[cell.rowi,cell.coli].should == cell.value
    end
  end
  it 'can save file to file and the content is the same as when saving to file' do
    book = Rspreadsheet.new($test_filename)    # open test file
    
    file = open(@tmp_filename, 'w') 
    book.save_to_io(file)              # and save the stream to @tmp_filename
    file.close

    book1 = Rspreadsheet.new($test_filename)   # now open both again
    book2 = Rspreadsheet.new(@tmp_filename)
    @sheet1 = book1.worksheets(1)
    @sheet2 = book2.worksheets(1)
    
    @sheet1.nonemptycells.each do |cell|       # and test if they are identical
      @sheet2[cell.rowi,cell.coli].should == cell.value
    end
  end
end




























