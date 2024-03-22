/*

*/

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("btn-format-font").onclick = () => tryCatch(setFontForAllSheets("Tahoma"));
    document.getElementById("btn-set-to-A1").onclick = () => tryCatch(setAllSheetsToA1);
  }
});

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

/**
 * Sets the font for all sheets in the current workbook.
 *
 * @param {string} fontName - The name of the font to set.
 */
function setFontForAllSheets(fontName) {
  // Get the current workbook
  Excel.run(function (context) {
    // Get all worksheets in the workbook
    var worksheets = context.workbook.worksheets;

    // Load worksheets
    worksheets.load("items");

    // Execute the request
    return context
      .sync()
      .then(function () {
        // Loop through each worksheet
        worksheets.items.forEach(function (worksheet) {
          // Get the used range of the worksheet
          var range = worksheet.getUsedRange();

          // Set the font for the entire used range
          range.format.font.name = fontName;
        });
      })
      .then(context.sync);
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

/**
 * Sets the active cell to A1 on all sheets in the current workbook.
 */
function setAllSheetsToA1() {
  // Get the current workbook
  Excel.run(function (context) {
    // Get all worksheets in the workbook
    var worksheets = context.workbook.worksheets;

    // Load worksheets
    worksheets.load("items");

    // Execute the request
    return context.sync().then(function () {
      // Loop through each worksheet and set the active cell to A1
      worksheets.items.forEach(function (worksheet) {
        worksheet.activate(); // Activate the worksheet
        worksheet.getRange("A1").select(); // Set the active cell to A1
      });
    });
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
