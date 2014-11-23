# rspreadsheet

[![Build Status](https://api.travis-ci.org/gorn/rspreadsheet.svg?branch=master)](https://travis-ci.org/gorn/rspreadsheet) [![Coverage Status](https://coveralls.io/repos/gorn/rspreadsheet/badge.png)](https://coveralls.io/r/gorn/rspreadsheet)
 
Manipulating spreadsheets with Ruby. Read, modify, write or create new OpenDocument Spreadsheet files from ruby code.

## Contibutions, ideas and wishes welcomed

If you need any help or find a bug please [submit an issue](https://github.com/gorn/rspreadsheet/issues) here. I appreciate any feedback and even if you can not help with code, it is interesting for me to hear from you. Different people have different needs and I want to hear about them. If you are a programmer and you have any ideas, wishes, etc you are welcomed to fork the repository and submit a pull request preferably including a failing test.

Alhought this gem is still in alpha stage I use in my project it and it works fine. Currently I am experimenting with syntax to get stabilized. **Any suggestions regarding the syntax is very welcomed**
 
## Examples of usage
  
```ruby
require 'rspreadsheet'
book = Rspreadsheet.open('spreadsheet.ods')
sheet = book.worksheets[1]


# get value of a cell B5 (there are more ways to do this)
sheet.B5                       # => 'cell value'
sheet[5,2]                     # => 'cell value'
sheet.rows(5).cells(2).value   # => 'cell value'

# set value of a cell B5
sheet.F5 = 'text'
sheet[5,2] = 7
sheet.cells(5,2).value = 1.78

# working with cell format
sheet.cells(5,2).format.bold = true
sheet.cells(5,2).format.background = '#FF0000'

# calculating sum of cells in row
sheet.rows(5).cellvalues.sum
sheet.rows(5).cells.sum{ |cell| cell.value }

# iterating over list of people and displaying the data

total = 0
sheet.rows.each do |row|
  puts "Sponsor #{row[1]} with email #{row[2]} has donated #{row[3]} USD."
  total += row[3]
end
puts "Totally fundraised #{total} USD"

# saving file
book.save
book.save('different_filename.ods')

```

This is the basic functionality. However rspreadsheet should allows lots of alternative syntax, like described in [GUIDE.md](GUIDE.md)

## Installation

Add this line to your application's Gemfile:

    gem 'rspreadsheet'

And then execute (if this fails install bundle gem):

    $ bundle

Or install it yourself as:

    $ gem install rspreadsheet
    
gem is hosted on Rubygems - https://rubygems.org/gems/rspreadsheet

If you get this or similar error

    mkmf.rb can't find header files for ruby at /usr/lib/ruby/include/ruby.h 
    
then you might not have installed libxml for ruby. I.e. in debian something like <code>sudo aptitude install ruby-libxml</code> or using equivalent command in other package managers.

## Motivation and Ideas

This project arised from the necessity. Alhought it is not true that there are no ruby gems allowing to acess OpenDOcument spreadsheet, I did not find another decent one which would suit my needs. Most of them also look abandoned and inactive. I have investigated these options:

  * [ruby-ods](https://github.com/yalab/ruby-ods) - this one seems as if it never really started
  * [rodf](https://github.com/thiagoarrais/rodf)- this only server as builder, it can not read existing files
  * [rods](http://www.drbreinlinger.de/ruby/rods/) - this is pretty ok, but it has terrible syntax. I first thought of writing wrapper around it, but it turned to be not so easy. Also last commit is 2 years old.
  * [rubiod](https://github.com/netoctone/rubiod) - this one is quite ok, the syntax is definitely better that in rods, but it seems also very abandoned. It does not support formats. This is a closest match.
  * [spreadsheet](https://github.com/zdavatz/spreadsheet) - this does not work with OpenDocument and even with Excel has issues in modyfying document. However since it is supposedly used, and has quite good syntax it might be inspirative. I also find the way this gem handles lazy writing of new rows to Spreadsheet object flawed, as well as strange accesibility of rows array object, which if assigned breaks consistency of sheet.
  * [roo](https://github.com/roo-rb/roo) can only read spreadsheets and not modify and write them back.

One of the main ideas is that the manipulation with OpenDOcument files should be forward compatible and as much current data preserving as possible. The parts of the file which are not needed for the change should not be changed. This is different to some of the mentioned gems, which generate the document from scratch, therefore any advanced features present in the original file which are not directly supported are lost.

  
## Contributing

1. [Fork it](http://github.com/gorn/rspreadsheet/fork)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request

## Further reading

* [Advanced Guide](GUIDE.md) how to use the gem
* [Code documentation](http://www.rubydoc.info/github/gorn/rspreadsheet) is hosted on [rubydoc.info](http://www.rubydoc.info/)
* [Changelog](CHANGELOG.md)
* [Documentation for developers](DEVEL_BLOG.md) containing ideas for future development and documentation on testing tools
