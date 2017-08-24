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

      $.each(visible_columns, function(index, value){
        export_columns.push(value);
      })
      $.each(all_columns, function(index, value){
        if(value.excel_export == true){
          export_columns.push(value);
        }
      })

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
      var $excelModal = $('<div/>')
        .attr({'id': 'excel-modal'})
        .addClass('modal')
        .width('500px')
        .appendTo($('body'));
      var $modalContent = $('<div/>')
        .addClass('modal-content')
        .append($('<h5/>').text('Excel export'))
        .append($('<p/>').text('Please wait while your Excel document is being prepared.'))
        .append($('<div/>').addClass('progress').append($('<div/>').addClass('indeterminate')))
        .appendTo($excelModal);
      var $modalFooter = $('<div/>')
        .addClass('modal-footer')
        .append($('<div/>').addClass('btn modal-close').text('Close'))
        .appendTo($excelModal);

      $excelModal.modal({
        complete: function() {
          $excelModal.remove();
        }
      });

      $excelModal.modal('open');
    }

    function requestExcel() {
      var inputs = '<input type="hidden" name="columns" value="'+ $paramsUrl +'" />';
      // Add sorting, filters and params from the loader
      jQuery.each(($grid.query.replace('?', '') + "&" + $grid.loader.conditionalURI()).split('&'), function(){
       var pair = this.split('=');
       inputs += '<input type="hidden" name="'+ pair[0] +'" value="'+ decodeURIComponent(pair[1]) +'" />';
      });
      var $form =  $('<form action="'+ $grid.path +'.xlsx" method="GET" data-remote=true>' + inputs + '</form>');
      var screenName = $grid.screen;
      $form.on("ajax:success", function(data, status, xhr){
        $('#excel-modal').modal('close');
        $form.remove();
        if(status.file && status.name) {
          // another form for downloading
          var input_1 = '<input type="hidden" name="filepath" value="'+ status.file +'" />';
          var input_2 = '<input type="hidden" name="filename" value="'+ status.name +'" />';
          var input_3 = '<input type="hidden" name="screen" value="'+ screenName +'" />';
          $('<form action="'+ $grid.path +'.xlsx" method="GET">' + input_1 + input_2 + input_3 + '</form>').appendTo('body').submit().remove();
        }
        return false;
      })
      displayLoadingModal();
      $form.appendTo('body').submit();
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
