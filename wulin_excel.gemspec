# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "wulin_excel/version"

Gem::Specification.new do |s|
  s.name        = "wulin_excel"
  s.version     = WulinExcel::VERSION
  s.platform    = Gem::Platform::RUBY
  s.authors     = ["Ekohe"]
  s.email       = ["dev@ekohe.com"]
  s.homepage    = "http://rubygems.org/gems/wulin_excel"
  s.summary     = %q{Excel export support for WulinMaster}
  s.description = %q{Excel export support for WulinMaster}

  s.rubyforge_project = "wulin_excel"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
  
  s.add_dependency "write_xlsx"
end
