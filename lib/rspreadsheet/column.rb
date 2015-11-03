module Rspreadsheet

# Represents a columns in a spreadsheet. Similar to Row object, but 
# currently only partly implemented

class Column
  def initialize(aworksheet,acoli)
    @worksheet = aworksheet
    @coli = acoli
  end
  def cell(rowi)
    @worksheet.row(rowi).cell(@coli)
  
  end
end

end
