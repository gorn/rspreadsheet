# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'rspreadsheet/version'

Gem::Specification.new do |spec|
  spec.name          = "rspreadsheet"
  spec.version       = Rspreadsheet::VERSION
  spec.authors       = ["Jakub A.Těšínský"]
  spec.email         = ["jAkub.cz (A is at)"]
  spec.summary       = 'Manipulating LibreOffice Calc (OpenDocument Spreadsheet) files from Ruby (read / create / modify).'
  spec.description   = 'Manipulating LibreOffice Calc (OpenDocument Spreadsheet) files with Ruby. This gem can create new, read existing files abd modify them. When modyfying files, it tries to change as little as possible, making it as much forward compatible as possible.'

  spec.homepage      = "https://github.com/gorn/rspreadsheet"
  spec.license       = "GPL"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  def self.package_natively_installed?(pkgname)
    # if the shell fails, the system command returns nil. In that case we assume that the package is NOT installed. It might be overkill, because I am supresing the stderr as well
    (system("dpkg-query -l #{pkgname} 2>/dev/null | grep -q '^i'")==true) or  # debian based
    (system("rpm -qa 2>/dev/null | grep -q '#{pkgname}' ")==true) or          # rpm based 
    (system("pkg_info -q -e '#{pkgname}-*' >/dev/null 2>&1")==true)           # OpenBSD and alike (although OpenBSD registers the gem as a gem, so this is perhaps unnecessary
  end

  # runtime dependencies
  unless package_natively_installed?('ruby-libxml')
    spec.add_runtime_dependency 'libxml-ruby', '>=3.0'   # parsing XML files
  end
  spec.add_runtime_dependency 'rubyzip', '>=1.3'       # opening zip files
  spec.add_runtime_dependency 'andand', '>=1.3'

  # development dependencies

  spec.add_development_dependency "bundler", '>= 1.5', '<3'   # some 2.5 rubyies contain already bundler 2 included
  if Gem::Version.new(RUBY_VERSION) >= Gem::Version.new('2.4.5')
    # ruby 2.4.5 and up
    spec.add_development_dependency "rake", '>=10.0'
  else
    # ruby up to 2.4.5
    spec.add_development_dependency "rake", '~>0.9'
  end
  # testig - see http://bit.ly/1n5yM51
  spec.add_development_dependency "rspec", '>=2.0'       # running tests
  spec.add_development_dependency 'pry-nav', '>=0.0'     # enables pry 'next', 'step' commands
  spec.add_development_dependency "coveralls", '>=0.7' # inspecting coverage of tests


  # optional and testing
  if Gem::Version.new(RUBY_VERSION) >= Gem::Version.new('2.2.5')
#    spec.add_development_dependency "guard", '~>2.13'
#    spec.add_development_dependency "guard-rspec", '~>4.6'
  end

#   spec.add_development_dependency 'equivalent-xml'     # implementing xml diff

end

