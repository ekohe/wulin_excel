require "wulin_excel"
require "rails"

module WulinExcel
  class Engine < Rails::Engine
    engine_name :wulin_excel
    initializer "wulin_engine.assets", after: :append_assets_path, group: :all do |app|      
      if defined?(Propshaft)
        Rails.application.config.assets.paths << root.join("app", "assets", "stylesheets")
        Rails.application.config.assets.paths << root.join("app", "assets", "javascripts")
      end

      app.config.assets.precompile += %w( excel.js )
    end
  end
end
