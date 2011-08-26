require 'rubygems'
#require 'rails/all'
require 'rspec/rails'

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