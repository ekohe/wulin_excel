require 'wulin_excel/engine' if defined?(Rails)

module WulinExcel
  MAXIMUM_NUMBER_OF_ROWS = 65535
end

if !defined?(WulinMaster)
  raise "WulinExcel needs WulinMaster. Make sure WulinExcel is loaded after WulinMaster configuring config.plugins properly in application.rb"
end

# Add default export excel button to every grid
WulinMaster::Toolbar.add_to_default_toolbar "Excel", :class => 'excel_export', :icon => 'excel'

# Load excel javascript and stylesheet
WulinMaster::add_javascript 'excel.js'
WulinMaster::add_stylesheet 'excel.css'

# Register xls mime type
Mime::Type.register "application/vnd.ms-excel", :xls

require 'writeexcel'
require 'wulin_excel/action'