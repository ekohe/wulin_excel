require 'spec_helper'
require 'wulin_excel/action'

class CountriesController < ApplicationController
  attr_accessor :query
end

class Country < ActiveRecord::Base
end

describe CountriesController, :type => :controller do
  describe "Includes WulinMaster::Actions" do
    before :each do
      controller.class.send(:include, WulinMaster::Actions)
      Mime::Type.stub!(:lookup_by_extension).with("xls") {true}
    end
    
    it "should respond to render_xls if request format is xls" do
      controller.stub!(:params).and_return({:format => 'xls'})
      controller.should_receive(:render_xls)
      controller.index
    end
  end
end