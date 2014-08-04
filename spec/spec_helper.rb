RSpec.configure do |c|
  c.fail_fast = true
#   c.warnings = true
end

require 'coveralls'
Coveralls.wear!

$test_filename = './spec/testfile1.ods'

require 'rspreadsheet'
