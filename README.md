# WulinExcel

WulinExcel extends [WulinMaster](https://gitlab.ekohe.com/ekohe/wulin/wulin_master) with xlsx export. It adds an export button to the toolbar of every grid. The button requests the grid's current query and visible columns as an xlsx file and downloads it.

## Installation

Add the submodule at `vendor/gems/wulin_excel` on `develop`:

```bash
git submodule add -b develop git@gitlab.ekohe.com:ekohe/wulin/wulin_excel.git vendor/gems/wulin_excel
```

Add the gem to the `Gemfile`, after `wulin_master`:

```ruby
gem "wulin_excel", path: "vendor/gems/wulin_excel"
```

Import the script in `app/javascript/application.js`:

```js
import "../../vendor/gems/wulin_excel/app/javascript/wulin_excel.esm.js"
```

## Usage

The gem registers a default `export` action on `WulinMaster::Grid`, so every grid shows the export button. It also registers the `xls` and `xlsx` MIME types.

An export with more than `WulinExcel.maximum_number_of_rows` rows (default 65535) is refused with an error message. Change the limit in an initializer:

```ruby
WulinExcel.maximum_number_of_rows = 100_000
```

To add text to that error message, set `large_excel_warning` under `wulin_excel` in `APP_CONFIG`.

## Development

Ruby 4.0.9, Rails 8.1. `wulin_master` comes from its `v3` branch over SSH, so you need access to `gitlab.ekohe.com`.

```bash
bundle install
bundle exec rspec
gem build wulin_excel.gemspec
```

The specs run against an inline Rails app on SQLite (`spec/support/dummy_app.rb`).

## License

MIT, see `LICENSE`.
