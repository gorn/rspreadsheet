require 'xml'

ns = XML::Namespace.new(XML::Node.new('xxx'), 'soap', 'http://schemas.xmlsoap.org/soap/envelope/')


d = XML::Document.new
nr = XML::Node.new('root', 'text')
d.root = nr
n2 = XML::Node.new('element', 'content')
nr << n2
puts d.to_s

n2.namespaces.namespace = ns
nr.namespaces.namespace = ns
puts d.to_s

n2.namespaces.namespace = ns
puts d.to_s

