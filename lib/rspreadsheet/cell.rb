module Rspreadsheet
class Cell
  attr_reader :value
  def initialize()
    @value = nil
    #TODO: connect to xml node
  end
  def to_s
    value
  end
  def value=(avalue)
    @value=avalue
    self
  end
end
end
