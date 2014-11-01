# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'docx_mailmerge/version'

Gem::Specification.new do |spec|
  spec.name          = "docx_mailmerge"
  spec.version       = DocxMailmerge::VERSION
  spec.authors       = ["Gary Foster", "Stephen McIntosh"]
  spec.email         = ["gary@radicalbear.com", "stephen.mcintosh@radicalbear.com"]
  spec.summary       = "Performs a mail merge on DOCX (Microsoft Office Word) files"
  spec.description   = "Performs a mail merge on DOCX (Microsoft Office Word) files. The user can use the Mail Merge Manager feature of Microsoft Word to create templates with merge fields embedded in the document. This library then allows a developer to take that DOXC file template as an input along with a hash containing the merge field names and values, and will return a merged document as a DOCX file that can be saved as a file, or persisted in a paperclip attachment or any other programmatic use."
  spec.homepage      = "https://github.com/garyfoster/docx_mailmerge"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]
  
  spec.required_ruby_version = '>= 2.0.0'

  spec.add_dependency "rubyzip", "~> 1.1"
  spec.add_dependency "nokogiri", "~> 1.6"
  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake", "~> 10.1"
end
