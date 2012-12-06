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
      if (($grid==null) || ($grid.loader==null)) 
        return false;
      
      // Get columns order and size
      addColumnSizeAndOrder();

      // Request download
      requestExcel();
      
      return true;
    }
    
    function addColumnSizeAndOrder() {
      $paramsUrl = '';
      var colArray = [];
      $.each($grid.getColumns(), function(index, value) {
        colArray.push(value.column_name + '~' + value.width)
      });
      $paramsUrl += colArray.join();
    }
  
    function requestExcel() { 
      var inputs = '<input type="hidden" name="columns" value="'+ $paramsUrl +'" />';
      // Add sorting, filters and params from the loader
      jQuery.each(($grid.query.replace('?', '') + "&" + $grid.loader.conditionalURI()).split('&'), function(){ 
       var pair = this.split('=');
       inputs += '<input type="hidden" name="'+ pair[0] +'" value="'+ decodeURIComponent(pair[1]) +'" />'; 
      });
      $('<form action="'+ $grid.path +'.xlsx" method="GET">' + inputs + '</form>').appendTo('body').submit().remove();
    }
    
    init();

    return {"sendExcelRequest": sendExcelRequest};
  }
  $.extend(true, window, { Excel: Excel });
})(jQuery);


// Excel action
WulinMaster.actions.Excel = $.extend({}, WulinMaster.actions.BaseAction, {
  name: 'excel',

  handler: function() {
    var grid = this.getGrid();
    var excel = new Excel(grid.name);
    if (!excel.sendExcelRequest()) {
      displayErrorMessage("Excel generation failed. Please try again later.");
    }
    return false;  
  }
});

WulinMaster.ActionManager.register(WulinMaster.actions.Excel);

