module WulinExcel
  MAXIMUM_NUMBER_OF_ROWS = 65535
  
  # rails 3.1 specified
  class Engine < ::Rails::Engine
    config.before_configuration do   
      if !Object.const_defined?(:WulinMaster)
        raise "WulinExcel needs WulinMaster. Make sure WulinExcel is loaded after WulinMaster configuring config.plugins properly in application.rb"
      end
    end
  
    config.after_initialize do 
      # Add default export excel button to every grid
      WulinMaster::Grid.add_to_default_toolbar "Excel", :class => 'excel_export', :icon => 'excel'

      # Load excel javascript
      WulinMaster::add_javascript 'excel.js'

      # Register xls mime type
      Mime::Type.register "application/vnd.ms-excel", :xls
    
      require 'wulin_excel/action'
    end
  end

end