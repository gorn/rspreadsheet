# This tells Bundler only load gems in side gemspec (not the locally installed ones). See http://stackoverflow.com/questions/4398262/setup-rspec-to-test-a-gem-not-rails
require 'bundler/setup'
Bundler.setup

RSpec.configure do |config|
  config.fail_fast = true
#   config.warnings = true
  
  # so i can run individual test just by appending :focus to them
  config.treat_symbols_as_metadata_keys_with_true_values = true if config.respond_to? :treat_symbols_as_metadata_keys_with_true_values= # older form
#  config.filter_run_when_matching(:focus)     
    
## preparing for RSpec 4 (will be default in RSpec 4
  
  if config.respond_to? :expect_with and config.respond_to? :include_chain_clauses_in_custom_matcher_descriptions  
    config.expect_with :rspec do |expectations|
      expectations.include_chain_clauses_in_custom_matcher_descriptions = true
    end
  end
  config.mock_with :rspec do |mocks|
    mocks.verify_partial_doubles = true if mocks.respond_to? :verify_partial_doubles
  end
  config.shared_context_metadata_behavior = :apply_to_host_groups if config.respond_to? :shared_context_metadata_behavior

end

# this enables Coveralls
require 'coveralls'
Coveralls.wear!

# some variables used everywhere
$test_filename = './spec/testfile1.ods'
$test_filename_images = './spec/testfile2-images.ods'

# require my gem
require 'rspreadsheet'



