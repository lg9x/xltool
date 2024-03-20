import * as common from "./common.js";

export async function updateWBS() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const sheet = context.workbook.worksheets.getItem("WBS");
      var table = sheet.tables.getItem("TblWbs");

      var rangeTable = table.getRange();
      rangeTable.load("values");
      await context.sync();
      var headerTable = rangeTable.values[0];
      //TODO: write a function that get column index
      var p_effort_columnIndex = common.getColumnIndexByHeaderName(headerTable, "P.Effort");
      var a_effort_columnIndex = common.getColumnIndexByHeaderName(headerTable, "A.Effort");
      var p_completed_columnIndex = common.getColumnIndexByHeaderName(headerTable, "P. % Completed");

      for (let i = 0; i < rangeTable.values.length; i++) {
        var p_effort = rangeTable.values[i][p_effort_columnIndex];
        var a_effort = rangeTable.values[i][a_effort_columnIndex];
        var p_completed = a_effort / p_effort;
        var cellRange = rangeTable.getCell(i, p_completed_columnIndex);
        cellRange.values = [[p_completed]];
        cellRange.numberFormat = [["0.00%"]];
      }
    });
  } catch (err) {
    console.error(err);
  }
}

function getRowByWBSId(rangeTable, wbsId) {
  for (let i = 0; i < rangeTable.values.length; i++) {
    if (rangeTable.values[i][0] == wbsId) {
      return rangeTable.values[i];
    }
  }
}
