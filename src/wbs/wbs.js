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
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */

      writeHelloToToolCheckColumn(context);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

function writeHelloToToolCheckColumn(context) {
  // Lấy worksheet hiện tại
  var sheet = context.workbook.worksheets.getActiveWorksheet();

  // Lấy bảng theo tên
  var table = sheet.tables.getItem("Tbl_WBS");
  var toolCheckColumn = table.columns.getItem("Tool check");

  // Lấy toàn bộ phạm vi của cột (bỏ qua tiêu đề)
  var toolCheckColumnRange = toolCheckColumn.getDataBodyRange();

  // Đặt giá trị "Hello" cho toàn bộ cột
  toolCheckColumnRange.values = [["Hello"]];
}
