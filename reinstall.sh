#!/bin/sh

cd rspreadsheet
gem build rspreadsheet.gemspec
sudo gem install rspreadsheet-0.1.*.gem
cd ..
