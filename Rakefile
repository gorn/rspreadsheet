require "bundler/gem_tasks"
require "rspec/core/rake_task"

RSpec::Core::RakeTask.new(:spec)  do |task|
  task.rspec_opts = ['--color', '--format', 'd']  # jak se ma formatova vystup
end

task :default => :spec
