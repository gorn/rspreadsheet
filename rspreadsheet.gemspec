# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'rspreadsheet/version'

Gem::Specification.new do |spec|
  spec.name          = "rspreadsheet"
  spec.version       = Rspreadsheet::VERSION
  spec.authors       = ["Jakub A.Těšínský"]
  spec.email         = ["jAkub.cz (A is at)"]
  spec.summary       = %q{Manipulating spreadsheets with Ruby (read / create / modify OpenDocument Spreadsheet).}
  spec.description   = %q{Manipulating OpenDocument spreadsheets with Ruby. This gem can create new, read existing files abd modify them. When modyfying files, it tries to change as little as possible, making it as much forward compatifle as possible.}
  spec.homepage      = "https://github.com/gorn/rspreadsheet"
  spec.license       = "GPL"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  def self.package_installed?(pkgname)
    Gem::Specification.find_all_by_name(pkgname).length > 0
  end
  
  # runtime dependencies
  unless package_installed?('ruby-libxml')
    spec.add_runtime_dependency 'libxml-ruby', '~>2.7'   # parsing XML files
  end
  spec.add_runtime_dependency 'rubyzip', '~>1.1'       # opening zip files
  spec.add_runtime_dependency 'andand', '~>1.3'

  # development dependencies
  spec.add_development_dependency "bundler", '~> 1.5'
  spec.add_development_dependency "rake", '~>0.9'
  # testig - see http://bit.ly/1n5yM51
  spec.add_development_dependency "rspec", '~>2.0'       # running tests
  spec.add_development_dependency 'pry-nav', '~>0.0'     # enables pry 'next', 'step' commands
  spec.add_development_dependency "coveralls", '~>0.7' # inspecting coverage of tests

  # optional and testing
  if Gem::Version.new(RUBY_VERSION) >= Gem::Version.new('2.2.5')
    spec.add_development_dependency "guard", '~>2.13'
    spec.add_development_dependency "guard-rspec", '~>4.6'
  end
  
#   spec.add_development_dependency 'equivalent-xml'     # implementing xml diff

end

