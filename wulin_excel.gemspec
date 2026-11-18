require_relative "lib/wulin_excel/version"

Gem::Specification.new do |s|
  s.name = "wulin_excel"
  s.version = WulinExcel::VERSION
  s.authors = ["Ekohe"]
  s.email = ["dev@ekohe.com"]
  s.homepage = "https://gitlab.ekohe.com/ekohe/wulin/wulin_excel"
  s.summary = "Excel export support for WulinMaster"
  s.description = "Adds an export button to WulinMaster grids that downloads the grid as an xlsx file."
  s.license = "MIT"

  s.files = `git ls-files`.split("\n")
  s.require_paths = ["lib"]

  s.required_ruby_version = ">= 3.4"

  s.add_dependency "rails", ">= 8.1"
  s.add_dependency "write_xlsx"

  s.add_development_dependency "rspec-rails"
  s.add_development_dependency "sqlite3"
  s.add_development_dependency "propshaft"
  s.add_development_dependency "dartsass-rails"
  s.add_development_dependency "roo"
  s.add_development_dependency "ostruct"
end
