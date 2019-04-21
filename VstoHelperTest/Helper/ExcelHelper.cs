using Excel = Microsoft.Office.Interop.Excel;

namespace VstoHelperTest.Helper
{
    public static class ExcelHelper
    {
        //**********************************************
        //worksheet handling
        //**********************************************

        public static int GetUsedRow(Excel.Worksheet wsTarget, int targetColumnPosition)
        {
            return ((Excel.Range)wsTarget.Cells[wsTarget.Rows.Count, targetColumnPosition]).End[Excel.XlDirection.xlUp].Row;
        }

        public static int GetUsedCol(Excel.Worksheet wsTarget, int targetRowPosition)
        {
            return ((Excel.Range)wsTarget.Cells[targetRowPosition, wsTarget.Columns.Count]).End[Excel.XlDirection.xlToLeft].Column;
        }

        public static void CleanContents(Excel.Worksheet wsTarget)
        {
            wsTarget.UsedRange.Clear();
        }

        public static void CleanContents(Excel.Worksheet wsTarget, int startRow, int startCol)
        {
            var usedRow = GetUsedRow(wsTarget, startCol);
            var usedCol = GetUsedCol(wsTarget, startRow);
            wsTarget.Range[wsTarget.Cells[startRow, startCol], wsTarget.Cells[usedRow, usedCol]].Clear();
        }

        public static void PasteArrayToSheet(dynamic arrContainer, Excel.Range rngTarget)
        {
            rngTarget.Resize[arrContainer.GetLength(0), arrContainer.GetLength(1)].Value = arrContainer;
        }

        public static void SaveAsFileFromSelectWorksheet(Excel.Worksheet wsTarget, string pathDetinationFolder, string nameTargetWorksheet)
        {
            var app = Globals.ThisWorkbook.Application;
            app.DisplayAlerts = false;
            wsTarget.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            wsTarget.Copy();
            app.Visible = true;
            var currentWorkbook = app.ActiveWorkbook;
            currentWorkbook.BuiltinDocumentProperties.Value = "INTERNAL";
            currentWorkbook.SaveAs(pathDetinationFolder + "\\" + nameTargetWorksheet + ".xlsx");
            app.DisplayAlerts = true;
        }

        public static void ConvertToExcelRealValue(Excel.Range selectRange)
        {
            selectRange.Value = selectRange.Value;
        }
    }
}
