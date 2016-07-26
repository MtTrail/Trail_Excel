# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'trail_excel/version'

Gem::Specification.new do |spec|
  spec.name          = "trail_excel"
  spec.version       = TrailExcel::VERSION
  spec.authors       = ["Mt.Trail"]
  spec.email         = ["trail@trail4you.com"]

  spec.summary       = %q{Control EXCEL from ruby script}
  spec.description   = %q{Control EXCEL from ruby script}
  spec.homepage      = "http://www.trail4you.com/TechNote/Ruby/Trail_Selenium.doc/Worksheet.html"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.12"
  spec.add_development_dependency "rake", "~> 10.0"
end
