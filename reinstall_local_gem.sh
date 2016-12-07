#!/bin/sh

# remove previously build gems
sudo rm -f pkg/rspreadsheet-0.*.gem
sudo rm -f Gemfile.lock

# this is to update git index, because git lsfiles is used in .gemspec
git add .

# builds the gem to pkg directory and installs gem to local system
sudo rake install

# Note 1: If the last step fails with "mkmf.rb can't find header files for ruby at /usr/lib/ruby/include/ruby.h",
# you may want to ``sudo aptitude install ruby-dev``

