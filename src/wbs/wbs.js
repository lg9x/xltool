/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */

      //await writeHelloToToolCheckColumn(context);
      await checkAStartAndSetWarning();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function writeHelloToToolCheckColumn(context) {
  // Lấy worksheet hiện tại
  var sheet = context.workbook.worksheets.getActiveWorksheet();

  // Lấy bảng theo tên
  var table = sheet.tables.getItem("Tbl_WBS");
  const columns = table.columns.load("name");
  await context.sync();

  var notes = "";
  columns.items.forEach((column) => {
    if (column.name === "A.Start") {
      var columnRange = column.getDataBodyRange();
      columnRange.load("values");
      if (columnRange.values.length > 0) {
        notes = "";
      } else {
        notes = "Hello";
      }
    }
    if (column.name === "Tool check") {
      var columnRange = column.getDataBodyRange();
      columnRange.values = notes;
    }
  });
}

async function checkAStartAndSetWarning() {
  try {
    await Excel.run(async (context) => {
      // Lấy worksheet hiện tại
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Lấy bảng bằng tên
      const table = sheet.tables.getItem("Tbl_WBS");

      // Tải toàn bộ dữ liệu thân bảng và tên cột
      table.load("columns/name, rows/values");

      await context.sync();

      // Tìm chỉ số của cột "A.Start" và "Tool check"
      const pStartIndex = table.columns.items.findIndex((column) => column.name === "P.Start");
      const pEndIndex = table.columns.items.findIndex((column) => column.name === "P.End");
      const pEffortIndex = table.columns.items.findIndex((column) => column.name === "P.Effort");
      const pCompIndex = table.columns.items.findIndex((column) => column.name === "P.Completed");

      const aStartIndex = table.columns.items.findIndex((column) => column.name === "A.Start");
      const aEndIndex = table.columns.items.findIndex((column) => column.name === "A.End");
      const aEffortIndex = table.columns.items.findIndex((column) => column.name === "A.Effort");
      const aCompIndex = table.columns.items.findIndex((column) => column.name === "A.Completed");

      const sttIndex = table.columns.items.findIndex((column) => column.name === "Status");
      const toolCheckIndex = table.columns.items.findIndex((column) => column.name === "Tool check");

      if (pStartIndex === -1 || aStartIndex === -1 || toolCheckIndex === -1) {
        console.error('Column "A.Start" or "Tool check" not found.');
        return;
      }

      // Duyệt qua từng dòng và kiểm tra giá trị của "A.Start"
      table.rows.items.forEach((row, index) => {
        const pStartValue = row.values[0][pStartIndex];
        const pEndValue = row.values[0][pEndIndex];
        const pEffortValue = row.values[0][pEffortIndex];
        const pCompValue = row.values[0][pCompIndex];

        const aStartValue = row.values[0][aStartIndex];
        const aEndValue = row.values[0][aEndIndex];
        const aEffortValue = row.values[0][aEffortIndex];
        const aCompValue = row.values[0][aCompIndex];

        // check if aStartValue is blank then set value to Tool check column
        if (!aStartValue || aStartValue === "") {
          // Chỉ định giá trị "warning" cho cột "Tool check"
          table.rows.items[index].getRange().getCell(0, toolCheckIndex).values = [["warning"]];
        }

        // check if today is greater than planned start date then set color to "red" for actual start date
        if (new Date(pStartValue) < new Date() && (!aStartValue || aStartValue === "")) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, aStartIndex);
          cell.format.fill.color = "red";
        } else {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, aStartIndex);
          cell.format.fill.clear();
          cell.format.font.color = "black";
        }

        // check if today is greater than planned end date then set color to "red" for actual end date
        if (new Date(pEndValue) < new Date() && (!aEndValue || aEndValue === "")) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, aEndIndex);
          cell.format.fill.color = "red";
        } else {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, aEndIndex);
          cell.format.fill.clear();
          cell.format.font.color = "black";
        }

        // check if A.Effort is greater than P.Effort then set color to "red" else set color to "green"
        if (aEffortValue && aEffortValue !== "" && aEffortValue > pEffortValue) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, aEffortIndex);
          cell.format.fill.color = "red";
          cell.format.font.color = "white";
        } else {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, aEffortIndex);
          cell.format.fill.color = "green";
          cell.format.font.color = "white";
        }

        // update status based on the actual completed with planned completed
        if (aCompValue > pCompValue) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, sttIndex);
          cell.format.fill.color = "green";
          cell.format.font.color = "white";
          table.rows.items[index].getRange().getCell(0, sttIndex).values = [["ahead of scheduled"]];
        }
        if (aCompValue < pCompValue) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, sttIndex);
          cell.format.fill.color = "red";
          cell.format.font.color = "white";
          table.rows.items[index].getRange().getCell(0, sttIndex).values = [["behind schedule"]];
        }
        if (aCompValue === pCompValue) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, sttIndex);
          cell.format.fill.color = "green";
          cell.format.font.color = "white";
          table.rows.items[index].getRange().getCell(0, sttIndex).values = [["on schedule"]];
        }
        if (aCompValue === 1) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, sttIndex);
          cell.format.fill.color = "green";
          cell.format.font.color = "white";
          table.rows.items[index].getRange().getCell(0, sttIndex).values = [["Completed"]];
        }
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}

async function checkInputStart() {
  try {
    await Excel.run(async (context) => {
      // Lấy worksheet hiện tại
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Lấy bảng bằng tên
      const table = sheet.tables.getItem("Tbl_WBS");

      // Tải toàn bộ dữ liệu thân bảng và tên cột
      table.load("columns/name, rows/values");

      await context.sync();

      // Tìm chỉ số của cột "A.Start" và "Tool check"
      const pStartIndex = table.columns.items.findIndex((column) => column.name === "P.Start");
      const aStartIndex = table.columns.items.findIndex((column) => column.name === "A.Start");

      if (pStartIndex === -1 || aStartIndex === -1) {
        console.error('Column "P.Start" or "A.Start" not found.');
        return;
      }

      // Duyệt qua từng dòng và kiểm tra giá trị của "A.Start"
      table.rows.items.forEach((row, index) => {
        const pStartValue = new Date(row.values[0][pStartValue]);
        var today = new Date();
        const aStartValue = row.values[0][aStartIndex];
        // Nếu "A.Start" rỗng, set "warning" vào "Tool check"
        if (today >= pStartValue) {
          const cell = table.rows.getItemAt(index).getRange().getCell(0, columnIndex);
          // Đặt màu nền cho ô
          cell.format.fill.color = "red";
        }
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}
