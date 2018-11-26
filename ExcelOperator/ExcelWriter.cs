using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelOperator
{
    public enum Alignment
    {
        Left,
        Middle,
        Right
    }
    public class ExcelWriter
    {
        private Application excelApplication;
        private Workbooks excelWorkbooks;
        private Workbook excelWorkbook;

        public ExcelWriter()
        {
            excelApplication = new Application();
            excelWorkbooks = excelApplication.Workbooks;            
        }

        public void Create()
        {
            excelWorkbook = excelWorkbooks.Add(true);
        }

        public void AddSheets(params string[] sheetNames)
        {
            for (int i = sheetNames.Length - 1; i >= 0; i--)
            {
                Worksheet sheet = excelWorkbook.Worksheets.Add();
                sheet.Name = sheetNames[i];
            }

            foreach (Worksheet sheet in excelWorkbook.Worksheets)
            {
                if (sheet.Name == "sheet1" || sheet.Name == "Sheet1")
                {
                    sheet.Delete();
                    break;
                }
            }
        }

        private void UnitCells(Worksheet worksheet, int x1, int y1, int x2, int y2)
        {
            worksheet.Range[worksheet.Cells[x1, y1], worksheet.Cells[x2, y2]].Merge();
        }

        public void UnitCells(string sheetName, int x1, int y1, int x2, int y2)
        {
            var worksheet = GetSheet(sheetName);
            UnitCells(worksheet, x1, y1, x2, y2);
        }

        public void SetCellValue(string sheetName, int x, int y, object value)
        {
            GetSheet(sheetName).Cells[x, y] = value;
        }

        public Worksheet GetSheet(string name)
        {
            return excelWorkbook.Worksheets[name];
        }

        public void SetCellProperty(string sheetName, int x, int y, int size, string name, Alignment alignment, bool bold = false)
        {
            var sheet = GetSheet(sheetName);
            var range = sheet.Range[sheet.Cells[x, y], sheet.Cells[x, y]];
            range.Font.Name = name;
            range.Font.Size = size;
            range.Font.Bold = bold;
            switch (alignment)
            {
                case Alignment.Left:
                    range.HorizontalAlignment = Constants.xlLeft;
                    break;
                case Alignment.Middle:
                    range.HorizontalAlignment = Constants.xlCenter;
                    break;
                case Alignment.Right:
                    range.HorizontalAlignment = Constants.xlRight;
                    break;
            }

            SetCellsAutoSize(sheet);
        }

        private void SetCellsAutoSize(Worksheet sheet)
        {
            sheet.Cells.EntireColumn.AutoFit();
            sheet.Cells.EntireRow.AutoFit();
        }

        public void SetCellBorder(string sheetName, int x1, int y1, int x2, int y2)
        {
            var sheet = GetSheet(sheetName);
            sheet.Range[sheet.Cells[x1, y1], sheet.Cells[x2, y2]].Borders.LineStyle = XlLineStyle.xlContinuous;
            sheet.Range[sheet.Cells[x1, y1], sheet.Cells[x2, y2]].Borders.Weight = XlBorderWeight.xlThin;
        }

        public bool SaveAs(string fileName)
        {
            try
            {
                excelWorkbook.SaveAs(fileName, XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing);

                return true;
            }
            catch (Exception exception)
            {
                throw new Exception("保存失败!",exception);
            }
        }

        public void Close()
        {
            if (excelApplication == null) return;

            excelWorkbook?.Close(Type.Missing, Type.Missing, Type.Missing);
            excelWorkbooks?.Close();
            excelApplication.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
            excelWorkbook = null;
            excelWorkbooks = null;
            excelApplication = null;

            GC.Collect();
        }
    }
}
