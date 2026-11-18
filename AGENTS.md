# AGENTS.md

`wulin_excel` is a Rails engine for [`wulin_master`](https://gitlab.ekohe.com/ekohe/wulin/wulin_master). It adds xlsx export to every grid. Host apps load it as the `vendor/gems/wulin_excel` submodule: ceres, atlas, cruise, olatu, mima, mrl, bss, bss_ofe, waldo and ekohe-backoffice.

## Wiring it into a host app

Add the submodule, `gem "wulin_excel", path: "vendor/gems/wulin_excel"` after `wulin_master`, and import `vendor/gems/wulin_excel/app/javascript/wulin_excel.esm.js` in `app/javascript/application.js`. Steps are in the README.

## Gotchas

- `lib/wulin_excel.rb` raises when `WulinMaster` is not defined. Load the gem after `wulin_master`.
- `lib/wulin_excel/action.rb` reopens `WulinMaster::Actions#index` with `alias_method`. It handles `format=xlsx` and calls the original `index` for every other format. Change it together with `wulin_master`'s `actions.rb`.
- Download is two requests. The first builds the file in `Rails.root/tmp` and returns its path. The second passes `filepath` and `filename`, and the file is sent only when it sits directly in `Rails.root/tmp`.
- The row limit is `WulinExcel.maximum_number_of_rows` (default 65535). The app can add text to the error message with `APP_CONFIG['wulin_excel']['large_excel_warning']`.
- `vendor/assets/javascripts/excel.js` is the browser side. `wulin_excel.esm.js` imports it, so the host's esbuild bundles it. `excel.css.scss` and the sprockets precompile line in `engine.rb` are legacy; the icon path is used under Propshaft.

## Commands

```bash
bundle install
bundle exec rspec
gem build wulin_excel.gemspec
```

The specs run an inline Rails app on SQLite (`spec/support/dummy_app.rb`) with `wulin_master` from its `v3` branch. The Gemfile reads it over SSH.
