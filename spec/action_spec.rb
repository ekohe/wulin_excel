require 'spec_helper'

class CountriesController < ApplicationController
  attr_accessor :query
end

class Country < ActiveRecord::Base
end

describe CountriesController, :type => :controller do
  describe "Includes WulinMaster::Actions" do
    before :each do
      CountriesController.send(:include, WulinMaster::Actions)
    end
    
    it "should description" do
      get :index, :format => 'xls'
    end
  end
end