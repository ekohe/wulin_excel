require "wulin_excel"
require "rails"

module WulinExcel
  class Engine < Rails::Engine
    engine_name :wulin_excel
    initializer "add assets to precompile" do |app|
       app.config.assets.precompile += %w( excel.js excel.css excel_icon.png )
    end
  end
end
