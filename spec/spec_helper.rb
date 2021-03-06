require 'rubygems'
require 'spork'

Spork.prefork do
  ENV["RAILS_ENV"] ||= 'test'
  require 'rails/all'
  require 'rspec/rails'
  require 'writeexcel'
  
  RSpec.configure do |config|
    config.mock_with :rspec
  end
  
  # create a dummy rails app
  class TestApp < Rails::Application
    config.root = File.dirname(__FILE__)
  end
  Rails.application = TestApp
  
  TestApp.routes.draw do
    resources :countries
    root :to => "homepage#index"
  end
  
  module Rails
    def self.root
      @root ||= File.expand_path("../../tmp/rails", __FILE__)
    end
  end
  
  class ApplicationController < ActionController::Base
  end 
end

Spork.each_run do
  module WulinMaster
    module Actions
      def index
        "original index action"
      end
    end
  end
  
  Mime::Type.register "application/vnd.ms-excel", :xls
  Mime::Type.register "application/vnd.openxmlformats-offedocument.spreadsheetml.sheet", :xlsx
end