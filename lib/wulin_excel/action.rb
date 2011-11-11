module WulinMaster
  module Actions
    alias_method :index_without_excel, :index
    
    def index
      if params[:format].to_s == 'xls' and Mime::Type.lookup_by_extension("xls")
        render_xls
        return
      end
      index_without_excel      
    end  
    
    def render_xls      
      # Create initial query object
      @query = grid.model

      # Add the necessary where statements to the query
      construct_filters

      fire_callbacks :query_filters_ready

      # Always add a limit and offset
      @query = @query.limit(WulinExcel::MAXIMUM_NUMBER_OF_ROWS).offset(0)

      # Add order
      parse_ordering
      
      # Add select
      add_select
      
      fire_callbacks :query_ready

      # Get all the objects
      @objects = @query.all
      
      # start to build xls file
      filename = File.join(Rails.root, 'tmp', "export-#{ Time.now.strftime("%Y-%m-%d-at-%H-%M-%S") }.xls")
      workbook = WriteExcel.new(filename)
      worksheet  = workbook.add_worksheet
      
      columns = params[:columns].split(',').map{|x| x.split('-')}.map{|x| {'name' => x[0], 'width' => x[1]}}
      
      # build the header row for worksheet
      build_worksheet_header(workbook, worksheet, columns)
      
      # construct excel columns
      excel_columns = construct_excel_columns(columns)
      
      # build the content rows for worksheet
      build_worksheet_content(worksheet, @objects, excel_columns)
      
      # close the workbook and render file
      workbook.close
      send_data File.read(filename)
    end

  protected
  
    def construct_excel_columns(columns)
      excel_columns = []
      columns.each do |column|
        excel_columns << grid.columns.find{|col| col.name.to_s == column["name"].to_s }
      end
      excel_columns.compact # In case there's a column passed in the params[:column] that doesn't exist
    end
    
    def build_worksheet_header(book, sheet, columns)
      header_format = book.add_format
      header_format.set_bold
      header_format.set_align('top')
      sheet.set_row(0, 16, header_format)

      columns.each_with_index do |column, index|
        column_from_grid = grid.columns.find{|col| col.name.to_s == column["name"].to_s}
        label_text = column_from_grid.nil? ? column["name"] : column_from_grid.label
        sheet.write_string(0, index, label_text)
        sheet.set_column(index, index,  column["width"].to_i/6)
      end
    end
    
    def build_worksheet_content(sheet, objects, columns)
      i = 1
      objects.each do |object|
        j = 0
        sheet.set_row(i, 16)
        columns.each do |column|
          value = column.format(object.send(column.name.to_s)).to_s
          value.gsub!("\r", "") # Multiline fix
          sheet.write_string(i, j, value)
          j += 1
        end
        i += 1
      end
    end
  end
end