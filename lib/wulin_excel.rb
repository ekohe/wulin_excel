require 'wulin_excel/engine' if defined?(Rails)

module WulinExcel
  MAXIMUM_NUMBER_OF_ROWS = 65535
end

if !defined?(WulinMaster)
  raise "WulinExcel needs WulinMaster. Make sure WulinExcel is loaded after WulinMaster configuring config.plugins properly in application.rb"
end

# Add export button to grid as default toolbar item, choose default icon from https://material.io/icons/
WulinMaster::Grid.add_default_action :export, icon: :file_download, global: true

# Load excel javascript and stylesheet
WulinMaster::add_javascript 'excel.js'
WulinMaster::add_stylesheet 'excel.css'

# Register xls/xlsx mime type
Mime::Type.register "application/vnd.ms-excel", :xls
Mime::Type.register "application/vnd.openxmlformats-offedocument.spreadsheetml.sheet", :xlsx

# require 'writeexcel'
require 'write_xlsx'
require 'wulin_excel/action'
