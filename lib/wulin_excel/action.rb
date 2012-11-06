module WulinMaster
  module Actions
    alias_method :index_without_excel, :index
    
    def index
      if params[:format].to_s == 'xlsx' and Mime::Type.lookup_by_extension("xlsx")
        render_xlsx
        return
      end
      index_without_excel      
    end  
    
    def render_xlsx
      fire_callbacks :initialize_query

      # Create initial query object
      @query = @query || grid.model
      
      # Make sure the relation method is called to correctly initialize it
      # We had issues where it's not initialized through the relation method when using
      #  the where method
      grid.model.relation if grid.model.respond_to?(:relation)

      # Add the necessary where statements to the query
      construct_filters

      fire_callbacks :query_filters_ready

      # Always add a limit and offset
      @query = @query.limit(WulinExcel::MAXIMUM_NUMBER_OF_ROWS).offset(0)

      # Add order
      parse_ordering
      
      # Add includes (OUTER JOIN)
      add_includes
      
      # Add joins (INNER JOIN)
      add_joins
      
      fire_callbacks :query_ready

      # Get all the objects
      @objects = @query.all
      
      # Apply virtual attribute order
      apply_virtual_order
      
      # Apply virtual attribute filter
      apply_virtual_filter
      
      
      # start to build xls file
      filename = File.join(Rails.root, 'tmp', "export-#{ Time.now.strftime("%Y-%m-%d-at-%H-%M-%S") }.xlsx")
      workbook = WriteXLSX.new(filename)
      worksheet  = workbook.add_worksheet
      
      columns = params[:columns].split(',').map{|x| x.split('~')}.map{|x| {'name' => x[0], 'width' => x[1]}}
      
      # build the header row for worksheet
      build_worksheet_header(workbook, worksheet, columns)
      
      # construct excel columns
      excel_columns = construct_excel_columns(columns)
      
      # build the content rows for worksheet
      build_worksheet_content(worksheet, @objects, excel_columns)
      
      # close the workbook and render file
      workbook.close
      send_data File.read(filename), :filename => "#{grid.name}-#{Time.now.to_s(:db)}.xlsx"
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
        label_text = column_from_grid ? column_from_grid.label : column["name"]
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
          value = column.json(object)
          value = format_value(value, column)
          if Numeric === value
            sheet.write_number(i, j, value)
          else
            value = value.to_s
            value.gsub!("\r", "") # Multiline fix
            sheet.write_string(i, j, value)
          end
          j += 1
        end
        i += 1
      end
    end
    
    private
    
    def format_value(value, column)
      return value if String === value
      
      if Hash === value
        value[column.option_text_attribute].to_s
      elsif Array === value
        value.map{|x| x[column.option_text_attribute].to_s }.join(',')
      else
        value
      end
    end
    
  end
end