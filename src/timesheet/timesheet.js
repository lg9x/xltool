/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("add-data").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";
      range.values = [["Hello World"]];

      // add data to table timesheets
      // Example usage:
      var date = new Date(); // Date
      var pic = "John Doe"; // PIC
      var task = "Task A"; // Task
      var tow = "Department"; // TOW
      var hours = 8; // Hours
      var per_completed = 1; // Per Completed
      var description = "Worked on Task A"; // Description

      // Tạo một mảng với dữ liệu từ form
      var data = [
        [
          document.getElementById("date").value,
          document.getElementById("pic").value,
          document.getElementById("task").value,
          document.getElementById("tow").value,
          parseFloat(document.getElementById("hours").value),
          parseFloat(document.getElementById("percentComplete").value) / 100, // Chia cho 100 để chuyển thành giá trị phần trăm thực sự
          document.getElementById("description").value,
        ],
      ];

      // Remove the empty records from table timesheets
      deleteEmptyRowsFromTable();

      // Add data to Timesheet
      addDataToTimesheet(data);

      // format table body
      formatTableBodyWithNoColor();
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

function addDataToTimesheet(data) {
  return Excel.run(function (context) {
    // Get the active worksheet
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Find the table by name
    var table = sheet.tables.getItem("Tbl_Timesheet");

    // Add a new row to the table
    var newRow = table.rows.add(null /* add at the end */, data);

    // Synchronize the context to apply the changes
    return context.sync().then(function () {
      console.log("Data added to Timesheet table successfully.");
    });
  }).catch(function (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

/**
 * Xóa tất cả dòng trống trong bảng có tên là "Tbl_Timesheet", chỉ để lại tiêu đề.
 */
async function deleteEmptyRowsFromTable() {
  try {
    await Excel.run(async (context) => {
      // Lấy bảng dựa trên tên của nó
      var table = context.workbook.tables.getItem("Tbl_Timesheet");

      // Tải các dòng của bảng
      table.rows.load("values");

      // Chờ tải dữ liệu
      await context.sync();

      // Lưu trữ các index của dòng trống để xóa
      var emptyRowIndexes = [];

      // Duyệt qua các dòng của bảng để kiểm tra dòng trống
      for (var i = 0; i < table.rows.items.length; i++) {
        var isRowEmpty = table.rows.items[i].values[0].every((cell) => !cell || cell === "");
        if (isRowEmpty) {
          // Lưu index của dòng trống
          emptyRowIndexes.push(i);
        }
      }

      // Xóa dòng trống từ dưới lên trên để tránh thay đổi index của các dòng chưa được xem xét
      for (var i = emptyRowIndexes.length - 1; i >= 0; i--) {
        table.rows.getItemAt(emptyRowIndexes[i]).delete();
      }

      // Chờ đợi thực thi xóa
      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}

/**
 * Định dạng lại phần thân của bảng "Tbl_Timesheet" không có màu nền.
 */
async function formatTableBodyWithNoColor() {
  try {
    await Excel.run(async (context) => {
      // Lấy bảng dựa trên tên của nó
      var table = context.workbook.tables.getItem("Tbl_Timesheet");
      const columns = table.columns.load("name");
      await context.sync();

      // Lấy phần thân của bảng
      var bodyRange = table.getDataBodyRange();

      // Đặt màu nền của phần thân là không màu
      bodyRange.format.fill.clear();

      // Đặt màu của văn bản trong phần thân là màu đen
      bodyRange.format.font.color = "black";

      columns.items.forEach((column) => {
        const columnRange = column.getDataBodyRange();
        switch (column.name) {
          case "Date":
            columnRange.numberFormat = [["yyyy-mm-dd"]];
            break;
          case "Hours":
            columnRange.numberFormat = [["0"]];
            break;
          case "% completed":
            columnRange.numberFormat = [["0%"]];
            break;
        }
      });

      // Tự động điều chỉnh kích thước các cột dựa trên nội dung
      bodyRange.getEntireColumn().autofitColumns();

      // Chờ đợi thực thi
      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}
// Hàm hiển thị popup
function showPopup() {
  document.getElementById("popup").style.display = "block";
}

// Hàm đóng popup
function closePopup() {
  document.getElementById("popup").style.display = "none";
}
