guard 'rspec' do
  # watch /lib/rspreadsheet/ files
  watch(%r{^lib/rspreadsheet/(.+).rb$}) do |m|
    "spec/#{m[1]}_spec.rb"
  end

  # watch /spec/ files
  watch(%r{^spec/(.+).rb$}) do |m|
    "spec/#{m[1]}.rb"
  end
end