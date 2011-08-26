require 'spec_helper'
require 'wulin_excel/action'

class CountriesController < WulinMaster::ScreenController
  attr_accessor :query
end

class Country < ActiveRecord::Base
end

describe CountriesController, :type => :controller do
  describe "Includes WulinMaster::Actions" do
    before :each do
      
    end
  end
end