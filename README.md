# rspreadsheet

[![Build Status](https://api.travis-ci.org/gorn/rspreadsheet.svg?branch=master)](https://api.travis-ci.org/gorn/rspreadsheet) [![Coverage Status](https://coveralls.io/repos/gorn/rspreadsheet/badge.png)](https://coveralls.io/r/gorn/rspreadsheet)

Manipulating spreadsheets with Ruby. Read, modify, write or create new OpenDocument Spreadsheet files from ruby code.

## Contibutions, ideas and wishes welcomed

If you need any help or find a bug please [submit an issue](https://github.com/gorn/rspreadsheet/issues) here. I appreciate any feedback and even if you can not help with code, it is interesting for me to hear from you. Different people have different needs and I want to hear about them. If you are a programmer and you have any ideas, wishes, etc you are welcomed to fork the repository and submit a pull request preferably including a failing test.

**Beware: this gem is still in alpha stage (however I use it and it works for what I need).** Currently I am experimenting with syntax to get the feeling what could be the best.
 
## Examples of usage
  
```ruby
require 'rspreadsheet'

book = Rspreadsheet.open('./icecream_list.ods')
sheet = book.worksheets 'Icecream list'
total = 0

(3..20).each do |row|
  puts 'Icecream name: ' + sheet[row,2]
  puts 'Icecream ingredients: ' + sheet[row,3]
  puts "I ate this " + sheet[row,4] + ' times'
  total += sheet[row,4]
end

sheet[21,3] = 'Total:'
sheet[21,4] = total

sheet.rows[21].format.bold = true

book.save

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

One of the main ideas is that the manipulation with OpenDOcument files should be forward compatible and as much current data preserving as possible. The parts of the file which are not needed for the change should not be changed. This is different to some of the mentioned gems, which generate the document from scratch, therefore any advanced features present in the original file which are not directly supported are lost.

  
## Contributing

1. [Fork it](http://github.com/gorn/rspreadsheet/fork)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request

