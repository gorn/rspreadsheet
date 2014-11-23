def watch_all
  # watch /lib/rspreadsheet/ files
  watch(%r{^lib/rspreadsheet/(.+).rb$}) do |m|
    "spec/#{m[1]}_spec.rb"
  end

  # watch /spec/ files
  watch(%r{^spec/(.+).rb$}) do |m|
    "spec/#{m[1]}.rb"
  end
end

# classical part
scope group: :normal

group :normal do
  guard 'rspec' do watch_all end
end

# see http://stackoverflow.com/questions/18501471/guard-how-to-run-specific-tags-from-w-in-guards-console
group :focus do
  guard 'rspec', cli: '--tag focus' do watch_all end
end

#group :f do
#  guard 'rspec', cli: '--tag focus' do watch_all end
#end