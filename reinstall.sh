#!/bin/sh

cd rspreadsheet
gem build rspreadsheet.gemspec
gem install rspreadsheet-0.1.0.gem
cd ..
