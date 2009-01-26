require 'rubygems'
Gem::manage_gems
require 'rake/gempackagetask'

spec = Gem::Specification.new do |s|
	s.platform = Gem::Platform::RUBY
	s.name = 'activeshipping'
	s.version = "0.2"
	s.author = "Jonathan Bartlett"
	s.email = "jonathan@newmedio.com"
	s.summary = "Unified API for Shipping"
	s.files = FileList['lib/*.rb', 'README', 'lib/active_shipping/lib/*.rb', "GPL.txt", "LGPL.txt"].to_a
	s.require_path = 'lib'
	s.autorequire = 'activeshipping'
	s.has_rdoc = true
	s.extra_rdoc_files = ["README"]
end

Rake::GemPackageTask.new(spec) do |pkg|
	pkg.need_tar = true
end

task :default => "pkg/#{spec.name}-#{spec.version}.gem" do
	puts "Generated latest version"
end
