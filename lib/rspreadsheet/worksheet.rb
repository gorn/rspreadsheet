require 'rspreadsheet/row'

module Rspreadsheet
class Worksheet
  def initialize
    @rows = []
  end
  
  def cell(rowi,coli)
    row(rowi).cell(coli)
  end
  
  def row(rowi)
    @rows[rowi] || @rows[rowi] = Row.new(self,rowi)
  end

end
end
