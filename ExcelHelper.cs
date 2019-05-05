using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace RADHelper
{
    public class ExcelHelper
    {
        public int GetUsedRow(Excel.Worksheet wsTarget, int targetCol)
        {
            return ((Excel.Range)wsTarget.Cells[wsTarget.Rows.Count, targetCol])
                .End[Excel.XlDirection.xlUp].Row;
        }

        public int GetUsedCol(Excel.Worksheet wsTarget, int targetRow)
        {
            return ((Excel.Range)wsTarget.Cells[wsTarget.Columns.Count, targetRow])
                .End[Excel.XlDirection.xlToLeft].Column;
        }

        public void CleanContents(Excel.Worksheet wsTarget)
        {
            wsTarget.UsedRange.Clear();
        }

        public void CleanContents(Excel.Worksheet wsTarget
            , int startRow, int startCol)
        {
            var usedRow = GetUsedRow(wsTarget, startCol);
            var usedCol = GetUsedCol(wsTarget, startRow);

            ((Excel.Range)(wsTarget.Range[wsTarget.Cells[startRow, startCol]
                , wsTarget.Cells[startRow, startCol]])).Clear();
        }

        public static void PasteArrayDataToSheet(dynamic arrContainer
            , Excel.Range rngTarget)
        {
            rngTarget.Resize[arrContainer.GetLength(0)
                , arrContainer.GetLength(1)].Value = arrContainer;
        }

        public void ExportWorksheetAsExcel(Excel.Worksheet wsTarget, string targetFolderPath
            , string targetWsName, string docProperty = "INTERNAL")
        {
            var app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            //var app = Globals.ThisWorkbook.Application;
            app.DisplayAlerts = false;
            wsTarget.Copy();

            var currentWb = app.ActiveWorkbook;
            currentWb.BuiltinDocumentProperties("Comments").Value = "INTERNAL";
            currentWb.SaveAs($"{targetFolderPath}\\{targetWsName}.xlsx");
            currentWb.Close();
            app.DisplayAlerts = true;
        }

        public void ExportWorksheetAsPDF(Excel.Worksheet wsTarget, string targetFolderPath
            , string outputFileName)
        {
            //var app = Globals.ThisWorkbook.Application;
            var app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            app.DisplayAlerts = false;
            wsTarget.Copy();

            var currentWb = app.ActiveWorkbook;
            currentWb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF
                ,$"{targetFolderPath}\\{outputFileName}");
            currentWb.Close();
            app.DisplayAlerts = true;
        }

        public void ConvertToOriginalType(Excel.Range rngSelect)
        {
            rngSelect.Value = rngSelect.Value;
        }

        public Array ConvertRangeToObjectArray(Excel.Worksheet wsTarget
            , int startRow, int startCol)
        {
            var usedRow = GetUsedRow(wsTarget, startCol);
            var usedCol = GetUsedCol(wsTarget, startRow);
            var arrTemp = wsTarget.Range[wsTarget.Cells[startRow, startCol]
                , wsTarget.Cells[usedRow, usedCol]].Value;

            object[,] arrResult = new object[usedRow, usedCol];

            Array.Copy(arrTemp, 1, arrResult, 0, usedRow * usedCol);
            return arrResult;
        }

        public string[,] ConvertRangeToStringArray(Excel.Worksheet wsTarget
            , int startRow, int startCol)
        {
            var usedRow = GetUsedRow(wsTarget, startCol);
            var usedCol = GetUsedCol(wsTarget, startRow);
            var arrTemp = wsTarget.Range[wsTarget.Cells[startRow, startCol]
                , wsTarget.Cells[usedRow, usedCol]].Value;

            var dataHandler = new DataHandlingHelper();

            return dataHandler.ConverArrayToStringArray(arrTemp);
        }

        public string[,] ConvertRangeToStringArray(Excel.Worksheet wsTarget
            , int startRow, int startCol, int endRow, int endCol)
        {
            var arrTemp = wsTarget.Range[wsTarget.Cells[startRow, startCol]
                , wsTarget.Cells[endRow, endCol]].Value;

            var dataHandler = new DataHandlingHelper();

            return dataHandler.ConverArrayToStringArray(arrTemp);
        }

    }
}
