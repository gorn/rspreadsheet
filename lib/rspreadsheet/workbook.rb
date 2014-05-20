module Rspreadsheet
class Workbook
  def initialize
    @worksheets=[]
  end
  def create_worksheet
    @worksheets.push(Worksheet.new)
    return @worksheets.last
  end
  def open(filename)
    
  end
  def worksheet(ndx)
    @worksheets[ndx]
  end
end
end
