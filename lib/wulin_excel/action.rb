module WulinMaster
  module Actions
    alias_method :index_without_excel, :index

    def index
      if params[:format].to_s == 'xlsx' and Mime::Type.lookup_by_extension("xlsx")
        if params[:filepath] && params[:filename]
          send_file params[:filepath], :filename => params[:filename]
        else
          render_xlsx
        end
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

      fire_callbacks :query_initialized

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

      # If more than 2000 rows, cancel
      if @query.count > 2000
        render :js => "displayErrorMessage('Use pivot for large excel export')"
        return false
      end

      # Get all the objects
      @objects = @query.all unless @query.blank?

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
      build_worksheet_content(workbook, worksheet, @objects, excel_columns)

      # close the workbook and render file
      workbook.close
      render :json => {:file => filename, :name => "#{grid.name}-#{Time.now.to_s(:db)}.xlsx"}
    end

  protected

    def construct_excel_columns(columns)
      excel_columns = []
      columns.each do |column|
        excel_columns << grid.columns.find{|col| col.full_name == column["name"].to_s || col.name.to_s == column["name"].to_s }
      end
      excel_columns.compact # In case there's a column passed in the params[:column] that doesn't exist
    end

    def build_worksheet_header(book, sheet, columns)
      header_format = book.add_format
      header_format.set_bold
      header_format.set_align('top')
      sheet.set_row(0, 16, header_format)
      columns.each_with_index do |column, index|
        column_from_grid = grid.columns.find{|col| col.full_name == column["name"].to_s || col.name.to_s == column["name"].to_s}
        label_text = column_from_grid ? column_from_grid.label : column["name"]
        sheet.write_string(0, index, label_text)
        sheet.set_column(index, index,  column["width"].to_i/6)
      end
    end

    def build_worksheet_content(book, sheet, objects, columns)
      wrap_text_format = book.add_format
      wrap_text_format.set_text_wrap

      datetime_excel_formats = {}

      i = 1
      objects.each do |object|
        j = 0
        sheet.set_row(i)
        columns.each do |column|
          value = column.json(object)
          value = format_value(value, column)

          if Numeric === value
            sheet.write_number(i, j, value, wrap_text_format)
          elsif (not column.datetime_value.nil?)
            begin
              formatted_datetime = column.datetime_value.strftime("%Y-%m-%dT%H:%M:%S.000")
              datetime_excel_formats[column.datetime_excel_format] ||= book.add_format(:num_format => column.datetime_excel_format, :align => 'center')

              datetime_format = datetime_excel_formats[column.datetime_excel_format]

              sheet.write_date_time(i, j, formatted_datetime, datetime_format)
            rescue
              sheet.write_string(i, j, value.to_s, wrap_text_format)
            end
          else
            value = value.to_s
            value.gsub!("\r", "") # Multiline fix
            sheet.write_string(i, j, value, wrap_text_format)
          end
          j += 1
        end
        i += 1
      end if objects.present?
    end

    private

    def format_value(value, column)
      return value if String === value

      if Hash === value
        item = value[column.field_name]
        if item
          if Hash === item
            v = item[column.option_text_attribute]
            Numeric === v ? v : v.to_s
          elsif Array === item
            item.map{|x| x[column.option_text_attribute]}.join(",")
          end
        else
          value.inspect
        end
      else
        value
      end
    end

  end
end