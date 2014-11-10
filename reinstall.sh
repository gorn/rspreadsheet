#!/bin/sh

cd rspreadsheet
rm rspreadsheet-0.*.gem
gem build rspreadsheet.gemspec
sudo gem install rspreadsheet-0.*.gem
cd ..
