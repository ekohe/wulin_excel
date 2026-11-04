RSpec.describe "xlsx export", type: :request do
  let!(:france) { Country.create!(name: "France") }

  before do
    Person.create!(first_name: "Bob", age: 40, country: france, birthdate: Time.utc(1980, 5, 6, 7, 8))
    Person.create!(first_name: "Ann", age: 30, country: france, birthdate: Time.utc(1990, 1, 2, 3, 4))
    Person.create!(first_name: "Cid", age: 50, country: nil, birthdate: nil)
  end

  describe "creating the file" do
    it "answers with the file and download names as json" do
      export

      expect(response).to have_http_status(:ok)
      expect(response.media_type).to eq("application/json")
      body = JSON.parse(response.body)
      expect(body["file"]).to match(/\Aexport-.+\.xlsx\z/)
      expect(body["name"]).to match(/\Aperson-\d{4}-\d{2}-\d{2}-at-\d{2}-\d{2}-\d{2}\.xlsx\z/)
      expect(DUMMY_ROOT.join("tmp", body["file"])).to exist
    end

    it "writes a readable xlsx with the header from the grid column labels" do
      export(columns: "first_name~120,age~60,country~100")

      header = exported_rows(response).first
      expect(header).to eq(["First name", "Age", "Country"])
    end

    it "writes the columns in the order of the columns param" do
      export(columns: "age~60,first_name~120")

      expect(exported_rows(response).drop(1)).to eq([[30, "Ann"], [40, "Bob"], [50, "Cid"]])
    end

    it "orders the rows by the grid default sorting state" do
      export

      expect(exported_rows(response).drop(1).map(&:first)).to eq(%w[Ann Bob Cid])
    end

    it "orders the rows by sort_col and sort_dir" do
      export(sort_col: "people.age", sort_dir: "DESC")

      expect(exported_rows(response).drop(1).map(&:first)).to eq(%w[Cid Bob Ann])
    end

    it "keeps only the rows that match the filters" do
      export(filters: [{column: "age", value: "40", operator: "equals"}])

      expect(exported_rows(response).drop(1)).to eq([["Bob", 40]])
    end

    it "writes a header only sheet when no row matches" do
      export(filters: [{column: "age", value: "99", operator: "equals"}])

      expect(exported_rows(response).size).to eq(1)
    end

    it "writes nothing for a column in the columns param that the grid does not have" do
      export(columns: "first_name~120,missing~50")

      expect(exported_rows(response).drop(1)).to eq([["Ann", nil], ["Bob", nil], ["Cid", nil]])
    end

    it "accepts a column without width" do
      export(columns: "first_name")

      expect(exported_rows(response).drop(1).flatten).to eq(%w[Ann Bob Cid])
    end

    it "writes a sheet without columns for an empty columns param" do
      export(columns: "")

      expect(response).to have_http_status(:ok)
      expect(DUMMY_ROOT.join("tmp", JSON.parse(response.body)["file"])).to exist
    end
  end

  describe "association and nil values" do
    it "writes the association source, and an empty cell when the association is nil" do
      export(columns: "first_name~120,country~100")

      expect(exported_rows(response).drop(1)).to eq([["Ann", "France"], ["Bob", "France"], ["Cid", ""]])
    end

    it "writes dates as date cells and an empty cell for nil" do
      export(columns: "first_name~120,birthdate~100")

      rows = exported_rows(response).drop(1)
      expect(rows.map { |r| r[1].respond_to?(:strftime) ? r[1].strftime("%F %R") : r[1] }).to eq(["1990-01-02 03:04", "1980-05-06 07:08", ""])
    end

    it "writes numbers as numbers and booleans as text" do
      export(columns: "age~60,vip~40")

      expect(exported_rows(response).drop(1).first).to eq([30, "false"])
    end
  end

  describe "grid with its own multi-column select" do
    before { Post.create!(name: "p", person: Person.find_by!(first_name: "Ann")) }

    it "counts the rows without the custom select and exports the aggregate column" do
      get "/aggregates.xlsx", params: {columns: "first_name~120,posts_count~60"}, xhr: true

      expect(response).to have_http_status(:ok)
      expect(exported_rows(response).drop(1)).to eq([["Ann", 1], ["Bob", 0], ["Cid", 0]])
    end

    it "applies the row limit to that grid too" do
      WulinExcel.maximum_number_of_rows = 2
      get "/aggregates.xlsx", params: {columns: "first_name~120"}, xhr: true

      expect(response.body).to include("displayErrorMessage")
    end
  end

  describe "view without primary key that includes an association" do
    it "counts the rows and exports them" do
      get "/person_summaries.xlsx", params: {columns: "first_name~120,country~100"}, xhr: true

      expect(response).to have_http_status(:ok)
      expect(exported_rows(response).drop(1)).to eq([["Ann", "France"], ["Bob", "France"], ["Cid", ""]])
    end
  end

  describe "callback that swaps the query for another model" do
    it "counts on the query table and exports the header only" do
      get "/unfiltered_summaries.xlsx", params: {columns: "first_name~120"}, xhr: true

      expect(response).to have_http_status(:ok)
      expect(exported_rows(response)).to eq([["First name"]])
    end
  end

  describe "row limit" do
    it "answers with a javascript error message and writes no file" do
      WulinExcel.maximum_number_of_rows = 2
      export

      expect(response).to have_http_status(:ok)
      expect(response.media_type).to eq("text/javascript")
      expect(response.body).to eq("displayErrorMessage('The excel file is too large.');")
      expect(Dir[DUMMY_ROOT.join("tmp", "*")]).to be_empty
    end

    it "exports exactly the maximum number of rows" do
      WulinExcel.maximum_number_of_rows = 3
      export

      expect(exported_rows(response).size).to eq(4)
    end

    it "appends the large_excel_warning of APP_CONFIG" do
      stub_const("APP_CONFIG", {"wulin_excel" => {"large_excel_warning" => "Use filters."}})
      WulinExcel.maximum_number_of_rows = 2
      export

      expect(response.body).to eq("displayErrorMessage('The excel file is too large. Use filters.');")
    end

    it "defaults to 65535 rows" do
      expect(WulinExcel.maximum_number_of_rows).to eq(65535)
    end
  end

  describe "file name" do
    it "differs for two exports in the same second" do
      allow(Time).to receive(:now).and_return(Time.utc(2026, 11, 4, 10, 0, 0))

      export(filters: [{column: "age", value: "30", operator: "equals"}])
      first = JSON.parse(response.body)["file"]
      export(filters: [{column: "age", value: "40", operator: "equals"}])
      second = JSON.parse(response.body)["file"]

      expect(first).not_to eq(second)
    end

    it "keeps the rows of the first export when a second one runs in the same second" do
      allow(Time).to receive(:now).and_return(Time.utc(2026, 11, 4, 10, 0, 0))

      export(filters: [{column: "age", value: "30", operator: "equals"}])
      first = response
      export(filters: [{column: "age", value: "40", operator: "equals"}])

      expect(exported_rows(first).drop(1)).to eq([["Ann", 30]])
    end
  end

  describe "other formats and permissions" do
    it "leaves the json index untouched" do
      get "/people.json"

      expect(response).to have_http_status(:ok)
      expect(JSON.parse(response.body)["rows"].size).to eq(3)
    end

    it "does not export for a user without permission on the screen" do
      get "/forbiddens.xlsx", params: {columns: "first_name~120"}, xhr: true
    rescue ActionController::UnknownFormat
      expect(Dir[DUMMY_ROOT.join("tmp", "*")]).to be_empty
    else
      expect(response).to have_http_status(:unauthorized)
      expect(Dir[DUMMY_ROOT.join("tmp", "*")]).to be_empty
    end
  end
end
