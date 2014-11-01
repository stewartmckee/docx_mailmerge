# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'docx_mailmerge/version'

Gem::Specification.new do |spec|
  spec.name          = "docx_mailmerge"
  spec.version       = DocxMailmerge::VERSION
  spec.authors       = ["Gary Foster", "Stephen McIntosh"]
  spec.email         = ["gary@radicalbear.com", "stephen.mcintosh@radicalbear.com"]
  spec.summary       = "Performs a Mail Merge on docx (Microsoft Office Word) files"
  spec.description   = "Write a longer description. Optional."
  spec.homepage      = "https://github.com/garyfoster/docx_mailmerge"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency "rubyzip", "~> 1.1"
  spec.add_dependency "nokogiri"
  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake"
end
