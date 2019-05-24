# This tells Bundler only load gems in side gemspec (not the locally installed ones). See http://stackoverflow.com/questions/4398262/setup-rspec-to-test-a-gem-not-rails
require 'bundler/setup'
Bundler.setup

RSpec.configure do |config|
  config.fail_fast = true
#   config.warnings = true
  config.treat_symbols_as_metadata_keys_with_true_values = true # so i can run individual test just by appending :focus to them
  
end

# this enables Coveralls
require 'coveralls'
Coveralls.wear!

# some variables used everywhere
$test_filename = './spec/testfile1.ods'
$test_filename_images = './spec/testfile2-images.ods'

# require my gem
require 'rspreadsheet'
