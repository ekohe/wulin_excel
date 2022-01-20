// Excel Object definition
(function($) {
  function Excel(gridName) {
    var $grid;
    var $paramsUrl;

    function init() {
      $grid = gridManager.getGrid(gridName);
      $paramsUrl = '';
    }

    function sendExcelRequest() {
      if (($grid === null) || ($grid.loader === null))
        return false;

      // Get columns order and size
      addColumnSizeAndOrder();

      // Request download
      requestExcel();

      return true;
    }

    function addColumnSizeAndOrder() {
      $paramsUrl = '';
      var arr = {};
      var colArray = [];
      var visible_columns = $grid.getColumns();
      var all_columns = $grid.columns;
      var export_columns = [];

      const
        visible_names = visible_columns.map(col => col['column_name']),
        visible = col => visible_names.includes(col['column_name']),
        excelExportUndefined = col => col['excel_export'] === undefined || col['excel_export'] === null,
        excelExport = col => col['excel_export'] === true;
      all_columns.forEach(function(column) {
        if (visible(column)) {
          if (excelExport(column) || excelExportUndefined(column)) export_columns.push(column);
        } else {
          if (excelExport(column)) export_columns.push(column);
        }
      });

      $.each(export_columns, function(i, v){
        arr[export_columns[i]['id']] = export_columns[i];
      })
      export_columns = [];
      for(key in arr){
        export_columns.push(arr[key]);
      }

      $.each(export_columns, function(index, value) {
        colArray.push(value.id + '~' + value.width);
      });
      $paramsUrl += colArray.join();
    }

    function displayLoadingModal() {
      var $excelModal = Ui.baseModal()
        .attr({'id': 'excel-modal'})
        .width('500px');

      $excelModal.find('.modal-content')
        .append($('<h5/>').text('Excel export'))
        .append($('<p/>').text('Please wait while your Excel document is being prepared.'))
        .append($('<div/>').addClass('progress').append($('<div/>').addClass('indeterminate')))

      var $modalFooter = $('<div/>')
        .addClass('modal-footer')
        .append($('<div/>').addClass('btn modal-close').text('Close'))
        .appendTo($excelModal);
    }

    function requestExcel() {
      var screenName = $grid.screen;
      var path = $grid.path +'.xlsx?' + $grid.query.replace('?', '') + $grid.loader.conditionalURI() + '&' + "columns=" + $paramsUrl;

      $.ajax({
        url: path,
        type: 'GET'
      }).success(function(status) {
        $('#excel-modal').modal('close').remove();
        if(status.file && status.name) {
          var req = new XMLHttpRequest()
          req.open("GET", `${$grid.path}.xlsx?filepath=${status.file}&filename=${status.name}&screen=${screenName}`, true)
          req.responseType = "blob"

          req.onload = function (event) {
            var blob = req.response
            var link = document.createElement('a')
            link.href = window.URL.createObjectURL(blob)
            link.download = status.name
            link.click()

            // Remove the resource
            window.URL.revokeObjectURL(blob)
            link.remove()
          }
          req.send()
        }
        return false;
      })
      displayLoadingModal();
    }

    init();

    return {"sendExcelRequest": sendExcelRequest};
  }
  $.extend(true, window, { Excel: Excel });
})(jQuery);


// Export action
WulinMaster.actions.Export = $.extend({}, WulinMaster.actions.BaseAction, {
  name: 'export',

  handler: function() {
    var grid = this.getGrid();
    var excel = new Excel(grid.name);
    if (!excel.sendExcelRequest()) {
      displayErrorMessage("Excel generation failed. Please try again later.");
    }
    return false;
  }
});

WulinMaster.ActionManager.register(WulinMaster.actions.Export);
