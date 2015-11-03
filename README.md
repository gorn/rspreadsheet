
[![Build Status](https://api.travis-ci.org/gorn/rspreadsheet.svg?branch=master)](https://travis-ci.org/gorn/rspreadsheet) [![Coverage Status](https://coveralls.io/repos/gorn/rspreadsheet/badge.png)](https://coveralls.io/r/gorn/rspreadsheet) [![User opinion wanted](https://badge.waffle.io/gorn/rspreadsheet.png?label=user opinion wanted&title=User opinion wanted)](https://waffle.io/gorn/rspreadsheet)
# rspreadsheet
 
Manipulating spreadsheets with Ruby. Read, **modify**, write or create new OpenDocument Spreadsheet files from ruby code. 

The gem allows you to acces your file and modify any cell of it, **without** touching the rest of the file, which makes it compatible with all advanced features of ODS files (both existing and future ones). You do not have to worry if it supports feature XY, if it does not, it won't touch it. This itself makes it distinct from most of [similar gems](#why-another-opendocument-spreadsheet-gem). Alhought this gem is still in beta stage I use in everyday and it works fine.

## Examples of usage

```ruby
require 'rspreadsheet'
book = Rspreadsheet.open('./test.ods')
sheet = book.worksheets(1)

# get value of a cell B5 (there are more ways to do this)
sheet.B5                       # => 'cell value'
sheet[5,2]                     # => 'cell value'
sheet.row(5).cell(2).value   # => 'cell value'

# set value of a cell B5
sheet.F5 = 'text'
sheet[5,2] = 7
sheet.cell(5,2).value = 1.78

# working with cell format
sheet.cell(5,2).format.bold = true
sheet.cell(5,2).format.background_color = '#FF0000'

# calculate sum of cells in row
sheet.row(5).cellvalues.sum
sheet.row(5).cells.sum{ |cell| cell.value.to_f }

# or set formula to a cell
sheet.A1.formula='=SUM(A2:A9)'

# iterating over list of people and displaying the data
total = 0
sheet.rows.each do |row|
  puts "Sponsor #{row[1]} with email #{row[2]} has donated #{row[3]} USD."
  total += row[3].to_f
end
puts "Totally fundraised #{total} USD"

# saving file
book.save
book.save('different_filename.ods')
```

This is also pubished as Gist **where you can leave you comments and suggestions**:

  * [basic functionality](https://gist.github.com/gorn/42e33d086d9b4fda10ec) 
  * [extended examples](https://gist.github.com/gorn/b432e6a69e82628349e6) of lots of alternative syntax
  * [GUIDE.md](GUIDE.md) some other notes

## Installation and Configuration

Gem is on [Rubygems](https://rubygems.org/gems/rspreadsheet) so you can install it by

    $ gem install rspreadsheet

or add <code>gem 'rspreadsheet'</code> to your Gemfile and execute  <code>bundle</code>. 

If you get this error

    mkmf.rb can't find header files for ruby at /usr/lib/ruby/include/ruby.h 
    
then you might not have installed libxml for ruby. I.e. in debian something like <code>sudo aptitude install ruby-libxml</code> or using equivalent command in other package managers.

## Contibutions, ideas and wishes welcomed

### Nonprogrammers
If you need any help or find a bug please [submit an issue](https://github.com/gorn/rspreadsheet/issues) here. I appreciate any feedback and even if you can not help with code, it is interesting for me to hear from you. Different people have different needs and I want to hear about them. If you are a programmer and you have any ideas, wishes, etc you are welcomed to fork the repository and submit a pull request preferably including a failing test.

Alhought this gem is still in beta stage I use in everyday and it works fine. Currently I am experimenting with syntax to get stabilized. **Any suggestions regarding the syntax is very welcomed**

### Programmers

1. [Fork it](http://github.com/gorn/rspreadsheet/fork)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
 
## Why another OpenDocument spreadsheet gem?

I would be glad to safe myself work, but surprisingly, there are not that many gems for OpenDocument spreadsheets. Most of them also look abandoned and inactive, or can only read or write spreadsheets, but not modify them. I have investigated these options (you might as well):

  * [ruby-ods](https://github.com/yalab/ruby-ods) - this one seems abandoned, or even as if it never really started
  * [rodf](https://github.com/thiagoarrais/rodf)- this only serves as builder, it can not read existing files or modify them
  * [rods](http://www.drbreinlinger.de/ruby/rods/) - this is pretty ok, but it has terrible syntax. I first thought of writing wrapper around it, but it turned to be not so easy. Also last commit is 2 years old.
  * [rubiod](https://github.com/netoctone/rubiod) - this one is quite ok, the syntax is definitely better that in rods, but it seems also very abandoned. It does not support formats. This is a closest match.
  * [spreadsheet](https://github.com/zdavatz/spreadsheet) - this does not work with OpenDocument and even with Excel has issues in modyfying document. However since it is supposedly used, and has quite good syntax it might be inspirative. I also find the way this gem handles lazy writing of new rows to Spreadsheet object flawed, as well as strange accesibility of rows array object, which if assigned breaks consistency of sheet.
  * [roo](https://github.com/roo-rb/roo) can only read spreadsheets and not modify and write them back.

One of the main ideas is that the manipulation with OpenDOcument files should be forward compatible and as much current data preserving as possible. The parts of the file which are not needed for the change should not be changed. This is different to some of the mentioned gems, which generate the document from scratch, therefore any advanced features present in the original file which are not directly supported are lost.

## Further reading

* [Advanced Guide](GUIDE.md) how to use the gem
* [Configuration](CONFIGURATION.md) of the gem
* [Code documentation](http://www.rubydoc.info/github/gorn/rspreadsheet) is hosted on [rubydoc.info](http://www.rubydoc.info/)
* [Changelog](CHANGELOG.md)
* [Documentation for developers](DEVEL_BLOG.md) containing ideas for future development and documentation on testing tools
