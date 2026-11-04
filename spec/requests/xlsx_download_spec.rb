RSpec.describe "xlsx download", type: :request do
  let(:tmp) { DUMMY_ROOT.join("tmp") }

  def download(filepath, filename = "report.xlsx")
    get "/people.xlsx", params: {filepath: filepath, filename: filename}
  end

  before do
    Person.create!(first_name: "Ann", age: 30)
    FileUtils.mkdir_p(tmp.join("nested"))
    File.write(tmp.join("nested", "inner.xlsx"), "inner")
    File.write(DUMMY_ROOT.join("secret.txt"), "secret")
  end

  it "sends the file that the export created" do
    export
    file = JSON.parse(response.body)["file"]
    download(file, "People.xlsx")

    expect(response).to have_http_status(:ok)
    expect(response.body).to start_with("PK")
    expect(response.headers["Content-Disposition"]).to include('filename="People.xlsx"')
    expect(response.headers["Content-Type"]).to eq("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
  end

  it "sends back a file that Roo reads with the exported rows" do
    export
    download(JSON.parse(response.body)["file"])
    path = tmp.join("downloaded.xlsx")
    File.binwrite(path, response.body)

    expect(Roo::Excelx.new(path.to_s).row(2)).to eq(["Ann", 30])
  end

  it "answers 404 for a missing file" do
    expect { download("missing.xlsx") }.to raise_error(ActionController::MissingFile)
  end

  {
    "parent directory" => "../secret.txt",
    "parent and back" => "../tmp/nested/inner.xlsx",
    "absolute path" => "/etc/passwd",
    "nested directory" => "nested/inner.xlsx",
    "current directory" => "./nested/inner.xlsx",
    "empty path" => ""
  }.each do |label, filepath|
    it "answers 403 for #{label} (#{filepath.inspect})" do
      download(filepath)

      expect(response).to have_http_status(:forbidden)
      expect(response.body).to eq("Unauthorized")
    end
  end

  it "reads an absolute path relative to the tmp directory" do
    expect { download("/secret.txt") }.to raise_error(ActionController::MissingFile, /tmp\/secret\.txt/)
  end

  it "needs both filepath and filename, otherwise it exports" do
    get "/people.xlsx", params: {filepath: "x.xlsx", columns: "first_name~120"}, xhr: true

    expect(JSON.parse(response.body)).to have_key("file")
  end
end
