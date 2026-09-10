//= stub excel.esm
(function($) {
  function Excel(gridName) {
    const grid = gridManager.getGrid(gridName);
    let paramsUrl = "";

    function sendExcelRequest(options = {}) {
      if (!grid || !grid.loader) {
        return false;
      }

      // Get columns order and size
      collectColumnInformation();

      // Request download
      requestExcel(options);

      return true;
    }

    // Collects column information for export
    function collectColumnInformation() {
      const visibleColumns = grid.getColumns();
      const visibleColumnNames = visibleColumns.map(col => col["column_name"]);
      let exportColumns = [];

      // Filter columns for export
      visibleColumns.forEach(function(column) {
        const shouldExport = column["excel_export"] === true;
        const hasNoExportSetting = column["excel_export"] === undefined || column["excel_export"] === null;
        if (shouldExport || hasNoExportSetting) {
          exportColumns.push(column);
        }
      });

      // Remove duplicates
      const uniqueColumns = {};
      exportColumns.forEach(function(column) {
        uniqueColumns[column.id] = column;
      });

      exportColumns = Object.values(uniqueColumns);

      // Format column data for URL
      const columnPairs = exportColumns.map(column => `${column.id}~${column.width}`);
      paramsUrl = columnPairs.join(",");
    }

    // Displays a loading modal during export
    function displayLoadingModal() {
      const $excelModal = Ui.baseModal()
        .attr({
          "id": "excel-modal"
        })
        .width("500px");

      $excelModal.find(".modal-content")
        .append($("<h5/>").text("Excel export"))
        .append($("<p/>").text("Please wait while your Excel document is being prepared."))
        .append($("<div/>").addClass("progress").append($("<div/>").addClass("indeterminate")));

      $("<div/>")
        .addClass("modal-footer")
        .append($("<div/>").addClass("btn modal-close").text("Close"))
        .appendTo($excelModal);
    }

    // Makes the AJAX request to generate and download Excel file
    function requestExcel(options) {
      let columns = paramsUrl

      if (options["columns"]) {
        columns = options["columns"].join(",")
      }

      const screenName = grid.screen;
      const baseQuery = grid.query.replace("?", "");
      const conditionalParams = grid.loader.conditionalURI();
      const path = `${grid.path}.xlsx?${baseQuery}${conditionalParams}&columns=${columns}`;

      displayLoadingModal();

      $.ajax({
        url: path,
        type: "GET"
      }).success(function(status) {
        $("#excel-modal").modal("close").remove();

        if (status.file && status.name) {
          downloadFile(status.file, status.name, screenName);
        }
        return false;
      });
    }

    // Downloads the generated Excel file
    function downloadFile(filepath, filename, screenName) {
      const req = new XMLHttpRequest();
      req.open("GET", `${grid.path}.xlsx?filepath=${filepath}&filename=${filename}&screen=${screenName}`, true);
      req.responseType = "blob";

      req.onload = function() {
        const blob = req.response;
        const link = document.createElement("a");
        link.href = window.URL.createObjectURL(blob);
        link.download = filename;
        link.click();

        // Remove the resource
        window.URL.revokeObjectURL(blob);
        link.remove();
      };

      req.send();
    }

    return {
      sendExcelRequest
    };
  }

  // Add to global namespace
  $.extend(true, window, {
    Excel
  });
})(jQuery);

// Export action
WulinMaster.actions.Export = $.extend({}, WulinMaster.actions.BaseAction, {
  name: "export",
  handler: function() {
    const grid = this.getGrid();
    const excel = new Excel(grid.name);

    if (!excel.sendExcelRequest()) {
      displayErrorMessage("Excel generation failed. Please try again later.");
    }

    return false;
  }
});

// Register the Export action
WulinMaster.ActionManager.register(WulinMaster.actions.Export);