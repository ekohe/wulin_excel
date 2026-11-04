ENV["RAILS_ENV"] = "test"
ENV["DATABASE_URL"] = "sqlite3::memory:"

require "bundler/setup"
require "tmpdir"
require "fileutils"
require "rails/all"
require "propshaft"
require "dartsass-rails"
require "haml-rails"
require "wulin_master"
require "wulin_excel"
require "rspec/rails"
require "roo"
# wulin_master builds an OpenStruct for the unauthorized message without requiring it; apps get it from wulin_permits.
require "ostruct"

# Rails.root only provides the tmp directory that wulin_excel writes the xlsx files to.
DUMMY_ROOT = Pathname.new(Dir.mktmpdir("wulin_excel_dummy"))
FileUtils.mkdir_p(DUMMY_ROOT.join("tmp"))
at_exit { FileUtils.rm_rf(DUMMY_ROOT) }

module WulinExcelDummy
  class Application < Rails::Application
    config.root = DUMMY_ROOT
    config.eager_load = false
    config.secret_key_base = "wulin_excel_dummy"
    config.logger = Logger.new(nil)
    config.active_support.deprecation = :silence
    config.action_dispatch.show_exceptions = :none
    config.action_controller.allow_forgery_protection = false
    config.hosts.clear
  end
end

Rails.application.initialize!

require_relative "support/dummy_app"
require_relative "support/export_helpers"

RSpec.configure do |config|
  config.use_transactional_fixtures = true
  config.infer_spec_type_from_file_location!
  config.order = :random
  config.expect_with(:rspec) { |c| c.syntax = :expect }
  config.after { WulinExcel.maximum_number_of_rows = nil }
end
