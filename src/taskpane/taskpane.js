/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import * as wbs from "./wbs.js";
import * as common from "./common.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = () => tryCatch(run);
    document.getElementById("btn-02").onclick = () => tryCatch(writeVal);
    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
    document.getElementById("btn-create-wbs").onclick = () => tryCatch(createWBS);
    document.getElementById("btn-get-columnIndex").onclick = () =>
      tryCatch(getColumnIndexByHeaderName("Tbl_WBS", "Plan End"));
    document.getElementById("btn-set-a1").onclick = () => tryCatch(setAllSheetsToA1);
    document.getElementById("btn-set-font").onclick = () => tryCatch(setFontForAllSheets("Arial"));
    document.getElementById("btn-update-wbs").onclick = () => tryCatch(wbs.updateWBS);
    document.getElementById("btn-gen-sheets").onclick = () => tryCatch(genSheet);
  }
});

async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "red";

      await context.sync();
      const sheets = context.workbook.worksheets;
      sheets.getItem("Sheet1").getRange("A1").values = [["The selected address was " + range.address + "."]];
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function writeVal() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();
      range.values = "Xin chào tôi là Long đây";
    });
  } catch (error) {
    console.error(error);
  }
}

async function genSheet() {
  try {
    await Excel.run(async (context) => {
      var ws = context.workbook.worksheets;
      ws.add("Report");
      ws.add("WBS");
      ws.add("OT Plan");
      ws.add("Productivity");
      ws.add("MemberList");
      await context.sync();

      // add data
      createWBS();
    });

    await context.sync();
  } catch (error) {
    console.error(error);
  }
}

async function sampleFunction() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
    });

    await context.sync();
  } catch (error) {
    console.error(error);
  }
}

async function createWBS() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      var ws = context.workbook.worksheets;
      const wbs_ws = ws.getItem("WBS");
      const headerData = ["TaskID", "TaskName", "PIC", "% Complete", "Plan Start", "Plan End", "Plan Effort"];
      var rangeAddress = "A1:G1";
      const tbl_wbs = wbs_ws.tables.add(rangeAddress, true);
      tbl_wbs.name = "Tbl_WBS";

      tbl_wbs.getHeaderRowRange().values = [headerData];

      tbl_wbs.rows.add(null, [null]);
      tbl_wbs.getRange().format.autofitColumns();
      tbl_wbs.getRange().format.autofitRows();
      tbl_wbs.resize("A1:G50");
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function createTable() {
  await Excel.run(async (context) => {
    // TODO1: Queue table creation logic here.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    // TODO2: Queue commands to populate the table with data.
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    ]);
    // TODO3: Queue commands to format the table.
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

// Run this code within an Office.js context, such as inside an Office Add-in

// Function to get column index based on header name
function getColumnIndexByHeaderName(tableName, headerName) {
  // Get the current worksheet
  Excel.run(function (context) {
    // Get the specified table by name
    var table = context.workbook.tables.getItem(tableName);

    // Load the table columns
    table.columns.load("items");

    // Execute the request
    return context.sync().then(function () {
      // Find the column index based on the header name
      var columnIndex = -1;
      table.columns.items.forEach(function (column, index) {
        if (column.name === headerName) {
          columnIndex = index;
        }
      });

      if (columnIndex !== -1) {
        const range = context.workbook.getSelectedRange();
        range.values = "Column index for header '" + headerName + "': " + columnIndex;
        console.log("Column index for header '" + headerName + "': " + columnIndex);
      } else {
        console.log("Header '" + headerName + "' not found in table '" + tableName + "'.");
      }
    });
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// Run this code within an Office.js context, such as inside an Office Add-in

// Function to set the active cell of all sheets to A1
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

// Run this code within an Office.js context, such as inside an Office Add-in

// Function to set a specific font for all cells in all sheets
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

// Run this code within an Office.js context, such as inside an Office Add-in

// Function to upload and read data from a local file
function uploadAndReadFile() {
  Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /* 64 KB */ }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // File is uploaded successfully, proceed to read its content
      var file = result.value;
      var sliceCount = file.sliceCount;
      var slicesReceived = 0;
      var content = "";

      function processSlice(slice) {
        // Process each slice of the file
        content += slice.data;
        slicesReceived++;

        if (slicesReceived === sliceCount) {
          // All slices received, parse content and populate Excel sheet
          processData(content);
        } else {
          // Request the next slice
          file.getSliceAsync(slicesReceived);
        }
      }

      file.getSliceAsync(0, { asyncContext: { next: processSlice } }, processSlice);
    } else {
      console.log("Error uploading file: " + result.error.message);
    }
  });
}

// Function to process the content and populate Excel sheet
function processData(content) {
  // Parse the content, if necessary
  // For example, you can parse CSV or JSON data here

  // Populate an Excel sheet with the data
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    // Example: Writing content to cell A1
    sheet.getRange("A1").values = [[content]];
    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function test_sample(context) {}
