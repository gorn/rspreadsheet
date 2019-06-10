  require 'pry'

  module Rspreadsheet

# this module contains methods used bz several objects
module Tools
  using ClassExtensions if RUBY_VERSION > '2.1'
  
  def self.only_letters?(x); x.kind_of?(String) and x.match(/^[A-Za-z]*$/) != nil end
  def self.kind_of_integer?(x)
    (x.kind_of?(Numeric) and x.to_i==x) or 
    (x.kind_of?(String) and x.match(/^\d*(\.0*)?$/) != nil)
  end

  # converts cell adress like 'F12' to pair od integers [row,col]
  def self.convert_cell_address_to_coordinates(*addr)
    if addr.length == 1
      addr[0].to_s.match(/^([A-Za-z]{1,3})(\d{1,8})$/)
      colname = $~[1]
      rowi = $~[2].to_i
    elsif addr.length == 2
      a = addr[0]; b = addr[1]
      if a.kind_of?(Integer) and b.kind_of?(Integer) # most common case first
        colname,rowi = b,a 
      elsif only_letters?(a)
        if kind_of_integer?(b)
          colname,rowi = a,b.to_i
        else
          raise 'Wrong parameters - first is letters, but the seconds is not digits only'
        end
      elsif kind_of_integer?(a)
        if only_letters?(b)
          colname,rowi = b,a.to_i
        elsif kind_of_integer?(b)
          colname,rowi = b.to_i,a.to_i
        else
          raise 'Wrong second out of two paremeters - mix of digits and numbers'
        end
      else 
        raise 'Wrong first out of two paremeters - mix of digits and numbers'
      end
    else
      raise 'Wrong number of arguments'
    end
     
    return [rowi,convert_column_name_to_index(colname)]
  end
  
  def self.convert_column_name_to_index(colname)
    return colname if colname.kind_of?(Integer)
    ## first possibility how to implement it
    # colname=colname.rjust(3,'@')
    # return (colname[-1].ord-64)+(colname[-2].ord-64)*26+(colname[-3].ord-64)*26*26

    ## second possibility how to implement it
    # return (colname.to_i(36)-('A'*colname.size).to_i(36)).to_s(36).to_i(26)+('1'*colname.size).to_i(26)
    
    ## third possibility how to implement it (second one little shortened)
    s=colname.size
    return (colname.to_s.upcase.to_i(36)-(36**s-1).div(3.5)).to_s(36).to_i(26)+(26**s-1)/25
  end
  
  def self.convert_cell_coordinates_to_address(*coords)
    coords = coords[0] if coords.length == 1
    raise 'Wrong number of arguments' if coords.length != 2
    row = coords[0].to_i # security against string arguments
    col = coords[1].to_i
    colstring = ''
    if col > 702
      pom = (col-703).div(26*26)+1
      colstring += (pom+64).chr
      col -= pom*26*26
    end
    if col > 26
      pom = (col-27).div(26)+1
      colstring += (pom+64).chr
      col -= pom*26
    end
    colstring += (col+64).chr
    return colstring+row.to_s
  end
  def self.c2a(*x); convert_cell_coordinates_to_address(*x) end
  def self.a2c(*x); convert_cell_address_to_coordinates(*x) end
  # Finds {LibXML::XML::Namespace} object by its prefix. It knows all OpenDocument commonly used namespaces.
  # @return [LibXML::XML::Namespace] namespace object
  # @param prefix [String] namespace prefix
  def self.get_namespace(prefix)
    ns_array = {
      'office'=>"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
      'style'=>"urn:oasis:names:tc:opendocument:xmlns:style:1.0",
      'text'=>"urn:oasis:names:tc:opendocument:xmlns:text:1.0",
      'table'=>"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
      'draw'=>"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
      'fo'=>"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
      'xlink'=>"http://www.w3.org/1999/xlink",
      'dc'=>"http://purl.org/dc/elements/1.1/",
      'meta'=>"urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
      'number'=>"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
      'presentation'=>"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
      'svg'=>"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
      'chart'=>"urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
      'dr3d'=>"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
      'math'=>"http://www.w3.org/1998/Math/MathML",
      'form'=>"urn:oasis:names:tc:opendocument:xmlns:form:1.0",
      'script'=>"urn:oasis:names:tc:opendocument:xmlns:script:1.0",
      'ooo'=>"http://openoffice.org/2004/office",
      'ooow'=>"http://openoffice.org/2004/writer",
      'oooc'=>"http://openoffice.org/2004/calc",
      'dom'=>"http://www.w3.org/2001/xml-events",
      'xforms'=>"http://www.w3.org/2002/xforms",
      'xsd'=>"http://www.w3.org/2001/XMLSchema",
      'xsi'=>"http://www.w3.org/2001/XMLSchema-instance",
      'rpt'=>"http://openoffice.org/2005/report",
      'of'=>"urn:oasis:names:tc:opendocument:xmlns:of:1.2",
      'xhtml'=>"http://www.w3.org/1999/xhtml",
      'grddl'=>"http://www.w3.org/2003/g/data-view#",
      'tableooo'=>"http://openoffice.org/2009/table",
      'drawooo'=>"http://openoffice.org/2010/draw",
      'calcext'=>"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
      'loext'=>"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
      'field'=>"urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
      'formx'=>"urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
      'css3t'=>"http://www.w3.org/TR/css3-text/",
      'manifest'=>"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    }
    if !defined?(@pomnode) or @pomnode.nil?
      @pomnode = LibXML::XML::Node.new('xxx')
    end
    if !defined?(@ns) or @ns.nil? then @ns={} end
    if @ns[prefix].nil?
      @ns[prefix] = LibXML::XML::Namespace.new(@pomnode, prefix, ns_array[prefix])
    end
    return @ns[prefix]
  end
  # sets namespaced attribute "ns_prefix:key" in node to value. if value == delete_value then remove the attribute
  def self.set_ns_attribute(node,ns_prefix,key,value,delete_value=nil)
    raise 'Tools.set_ns_attribute can not set attribute on nil node' if node.nil?
    ns = Tools.get_namespace(ns_prefix)
    attr = node.attributes.get_attribute_ns(ns.href, key)
    
    unless value==delete_value # set attribute
      if attr.nil? # create attribute if needed
        attr = LibXML::XML::Attr.new(node, key,'temporarilyempty')
        attr.namespaces.namespace = ns
      end
      attr.value = value.to_s
      attr
    else # remove attribute
      attr.remove! unless attr.nil? 
      nil
    end
  end
  def self.get_ns_attribute(node,ns_prefix,key,default=:undefined_default)
    if default==:undefined_default
      raise 'Nil does not have any attributes' if node.nil?
      node.attributes.get_attribute_ns(Tools.get_namespace(ns_prefix).href,key)
    else
      node.nil? ? default : node.attributes.get_attribute_ns(Tools.get_namespace(ns_prefix).href,key) || default
    end
  end
  def self.get_ns_attribute_value(node,ns_prefix,key,default=:undefined_default)
    if default==:undefined_default
      Tools.get_ns_attribute(node,ns_prefix,key).andand.value
    else
      node.nil? ? default : Tools.get_ns_attribute(node,ns_prefix,key,nil).andand.value || default
    end
  end
  def self.remove_ns_attribute(node,ns_prefix,key)
    ns = Tools.get_namespace(ns_prefix)
    attr = node.attributes.get_attribute_ns(ns.href, key)
    attr.remove! unless attr.nil? 
  end
  def self.prepare_ns_node(ns_prefix,nodename,value=nil)
    LibXML::XML::Node.new(nodename,value, Tools.get_namespace(ns_prefix))
  end
  def self.insert_as_first_node_child(node,subnode)
    if node.first?
      node.first.prev = subnode
    else
      node << subnode
    end
  end
  
  def self.get_unused_filename(zip,prefix, extension)
    (1000..9999).each do |ndx|
      filename = prefix + ndx.to_s + ((Time.now.to_r*1000000000).to_i.to_s(16)) + extension
      return filename if zip.find_entry(filename).nil?
    end
    raise 'Could not get unused filename within sane times of iterations'
  end
 
  def self.new_time_value(h,m,s)
    Time.new(StartOfEpoch.year,StartOfEpoch.month,StartOfEpoch.day,h,m,s)
  end
  
  def self.output_to_zip_stream(io,&block)
    if io.kind_of? File or io.kind_of? String
      Zip::File.open(io, 'br+') do |zip|
        yield zip
      end
    elsif io.kind_of? StringIO # or io.kind_of? IO
      Zip::OutputStream.write_buffer(io) do |zip|
        yield zip
      end
    end
  end
 
  def self.content_xml_diff(filename1,filename2)
    content_xml1 = Zip::File.open(filename1) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    content_xml2 = Zip::File.open(filename2) do |zip|
      LibXML::XML::Document.io zip.get_input_stream('content.xml')
    end
    
    return xml_diff(content_xml1.root,content_xml2.root)
  end
  
  def self.xml_file_diff(filename1,filename2)
    content_xml1 = LibXML::XML::Document.file(filename1).root
    content_xml2 = LibXML::XML::Document.file(filename2).root
    return xml_diff(content_xml1, content_xml2)
  end

  def self.xml_diff(xml_node1,xml_node2)
    message = []
    message << xml_node2.first_diff(xml_node1)
    message << xml_node1.first_diff(xml_node2)
    message << 'content XML not equal' unless xml_node1.to_s.should == xml_node2.to_s
    message = message.compact.join('; ')
    message = nil if message == ''
    message
  end
end
  
end
