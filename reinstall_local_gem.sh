#!/bin/sh

# remove previously build gems
rm -f pkg/rspreadsheet-0.*.gem

# builds the gem to pkg directory and installs gem to local system
sudo rake install 

# Note 1: If the last step fails with "mkmf.rb can't find header files for ruby at /usr/lib/ruby/include/ruby.h",
# you may want to ``sudo aptitude install ruby-dev``

# Note 2: If the newest changes are not inclued in build, try to rm Gemfile.lock
