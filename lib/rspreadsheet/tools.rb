module Rspreadsheet

# this module contains methods used bz several objects
module Tools
  # converts cell adress like 'F12' to pair od integers [row,col]
  def self.convert_cell_address(*coords)
    if coords.length == 1
      coords[0].match(/^([A-Z]{1,3})(\d{1,8})$/)
      colname = $~[1]
      rowname = $~[2]
    elsif coords.length == 2
      colname = coords[0]
      rowname = coords[1]
    else
      raise 'Wrong number of arguments'
    end
      
    colname=colname.rjust(3,'@')
    col = (colname[-1].ord-64)+(colname[-2].ord-64)*26+(colname[-3].ord-64)*26*26
    row = rowname.to_i
    return [row,col]
  end
  
  # this object represents array which can contain repeated values
  # inspired valuely by http://www-users.cs.umn.edu/~saad/software/SPARSKIT/paper.ps 
  class SparseRepeatedArray < Array
    
  
  end
end
 
end