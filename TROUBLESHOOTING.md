# Troubleshooting

## Installation

### libxml

The gem detects natively installed libxml and uses instead of installing new one from rubygems. There might be some errors connected to this. If you get this error concering libxml, like

    checking for libxml/xmlversion.h in /opt/include/libxml2,/opt/local/include/libxml2,/usr/local/include/libxml2,/usr/include/libxml2... no
    *** extconf.rb failed ***
or 

    mkmf.rb can't find header files for ruby at /usr/lib/ruby/include/ruby.h 
        
then you might not have installed libxml for ruby. I.e. in debian something like 

    sudo aptitude install ruby-libxml
    
or using equivalent command in other package managers.

### if you do NOT want to use netively installed libxml

If you want to use libxml from rubygems, than uninstall the native one or add this line to your Gemfile

    'libxml-ruby', '3.0' 
