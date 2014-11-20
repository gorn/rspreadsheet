module Rspreadsheet

# this module contains methods used bz several objects
module Tools
  # converts cell adress like 'F12' to pair od integers [row,col]
  def self.convert_cell_address_to_coordinates(*addr)
    if addr.length == 1
      addr[0].match(/^([A-Z]{1,3})(\d{1,8})$/)
      colname = $~[1]
      rowname = $~[2]
    elsif addr.length == 2
      colname = addr[0]
      rowname = addr[1]
    else
      raise 'Wrong number of arguments'
    end
     
    ## first possibility how to implement it
#     colname=colname.rjust(3,'@')
#     col = (colname[-1].ord-64)+(colname[-2].ord-64)*26+(colname[-3].ord-64)*26*26

    ## second possibility how to implement it
    # col=(colname.to_i(36)-('A'*colname.size).to_i(36)).to_s(36).to_i(26)+('1'*colname.size).to_i(26)
    
    ## third possibility how to implement it (second one little shortened)
    s=colname.size
    col=(colname.to_i(36)-(36**s-1).div(3.5)).to_s(36).to_i(26)+(26**s-1)/25
    
    row = rowname.to_i
    return [row,col]
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
      'css3t'=>"http://www.w3.org/TR/css3-text/"
    }
    if @pomnode.nil?
      @pomnode = LibXML::XML::Node.new('xxx')
    end
    if @ns.nil? then @ns={} end
    if @ns[prefix].nil?
      @ns[prefix] = LibXML::XML::Namespace.new(@pomnode, prefix, ns_array[prefix])
    end
    return @ns[prefix]
  end
  # sets namespaced attribute "ns_prefix:key" in node to value. if value == delete_value then remove the attribute
  def self.set_ns_attribute(node,ns_prefix,key,value,delete_value=nil)
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
  def self.get_ns_attribute(node,ns_prefix,key)
    node.nil? ? nil : node.attributes.get_attribute_ns(Tools.get_namespace(ns_prefix).href,key)
  end
  def self.get_ns_attribute_value(node,ns_prefix,key)
    Tools.get_ns_attribute(node,ns_prefix,key).andand.value
  end
  def self.remove_ns_attribute(node,ns_prefix,key)
    node.attributes.get_attribute_ns(Tools.get_namespace(ns_prefix).href,key)
    attr.remove! unless attr.nil? 
  end
  def self.create_ns_node(ns_prefix,nodename,value=nil)
    LibXML::XML::Node.new(nodename,value, Tools.get_namespace(ns_prefix))
  end
end
 
end

# @private
class Range
  def size
    res = self.end-self.begin+1
    res>0 ? res : 0
  end
end