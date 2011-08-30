require 'spec_helper'
require 'wulin_excel/action'

class CountriesController < ApplicationController
  attr_accessor :query
end

class Country < ActiveRecord::Base
end

module WulinExcel
  MAXIMUM_NUMBER_OF_ROWS = 65535
end

describe CountriesController, :type => :controller do
  describe "Includes WulinMaster::Actions" do
    before :all do
      tmp_path = File.join(File.dirname(__FILE__), '..', 'tmp')
      FileUtils.mkdir tmp_path
      FileUtils.mkdir File.join(tmp_path, 'rails')
      FileUtils.mkdir File.join(tmp_path, 'rails', 'tmp')
    end
    
    after :all do
      tmp_path = File.join(File.dirname(__FILE__), '..', 'tmp')
      FileUtils.rm_rf tmp_path
    end
    
    before :each do
      controller.class.send(:include, WulinMaster::Actions)
      Mime::Type.stub!(:lookup_by_extension).with("xls") {true}
      
      @mock_query = mock("query")
      @mock_objects = mock("objects")
      
      @mock_query.stub!(:all) { @mock_objects }
      Country.stub_chain(:limit, :offset) { @mock_query }
      controller.stub_chain(:grid, :model) { Country }
           
      controller.stub!(:construct_filters) { true }
      controller.stub!(:fire_callbacks) { true }
      controller.stub!(:parse_ordering) { true }
      controller.stub!(:add_select) { true }
      controller.stub!(:build_worksheet_header) { true }
      controller.stub!(:construct_excel_columns) { true }
      controller.stub!(:build_worksheet_content) { true }
    end
    
    it "should call render_xls if request format is xls" do
      controller.stub!(:params).and_return({:format => 'xls'})
      controller.should_receive(:render_xls)
      controller.index
    end
    
    it "should not call render_xls if request format is not xls" do
      controller.stub!(:params).and_return({:format => "json"})
      controller.should_not_receive(:render_xls)
      controller.should_receive(:index_without_excel)
      controller.index
    end
    
    describe "render_xls" do
      before :each do
        controller.stub!(:send_data) { true }
        File.stub!(:read) { true }
        @mock_file = mock("file")
        @mock_file.stub!(:close) { true }
      end
      
      it "should render xls data" do
        controller.should_receive(:send_data)
        controller.render_xls
      end

      it "should only select WulinExcel::MAXIMUM_NUMBER_OF_ROWS defined rows" do
        Country.should_receive(:limit).with(65535)
        controller.render_xls
      end

      it "should create a xls file and create a xls worksheet, then close it" do
        @filename = "test.xls"
        @timestamp_str = "2011-08-30-at-18-30-00"
        Time.stub_chain(:now, :strftime) { @timestamp_str }
        
        File.should_receive(:join).with(Rails.root, 'tmp', "export-#{@timestamp_str}.xls").and_return(@filename)
        
        WriteExcel.should_receive(:new).with(@filename).and_return(@mock_file)
        @mock_file.should_receive(:add_worksheet)
        @mock_file.should_receive(:close)
        
        controller.render_xls
      end
    end
  end
end