module ExportHelpers
  def export(**params)
    get "/people.xlsx", params: {columns: "first_name~120,age~60"}.merge(params), xhr: true
  end

  # Reads the xlsx file that an export request created in Rails.root/tmp.
  def exported_sheet(response)
    Roo::Excelx.new(DUMMY_ROOT.join("tmp", JSON.parse(response.body)["file"]).to_s)
  end

  def exported_rows(response)
    sheet = exported_sheet(response)
    (sheet.first_row..sheet.last_row).map { |row| sheet.row(row) }
  end
end

RSpec.configure do |config|
  config.include ExportHelpers, type: :request
  config.before { FileUtils.rm_rf(Dir[DUMMY_ROOT.join("tmp", "*")]) }
end
