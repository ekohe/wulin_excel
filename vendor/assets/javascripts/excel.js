$(document).ready(function() {
  // Add click handler on the toolbar item
  $(".excel_export").live('click', function() {
    var grid_name = $(this).closest('.grid_container').attr('id').replace(/^grid_/, '');
    var excel = new Excel(grid_name);
    if (!excel.sendExcelRequest()) {
      displayErrorMessage("Excel generation failed. Please try again later.");
    }
    return false;      
  });
});

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
      
      // Add sorting, filters and params from the loader
      $paramsUrl += $grid.loader.conditionalURI();

      // Request download
      requestExcel();
      
      return true;
    }
    
    function addColumnSizeAndOrder() {
      $.each($grid.getColumns(), function(index, value) {
        $paramsUrl += "&columns[][name]="+value.id+"&columns[][width]="+value.width;
      });
    }
    
    function requestExcel() {
      var inputs = '';
  		jQuery.each($paramsUrl.split('&'), function(){ 
  			var pair = this.split('=');
  			inputs+='<input type="hidden" name="'+ pair[0] +'" value="'+ pair[1] +'" />'; 
  		});
      $('<form action="'+ $grid.path +'.xls" method="GET">'+inputs+'</form>').appendTo('body').submit().remove();
    }
    
    init();

		return {"sendExcelRequest": sendExcelRequest};
  }
	$.extend(true, window, { Excel: Excel });
})(jQuery);