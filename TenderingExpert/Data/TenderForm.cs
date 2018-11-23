using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelOperator;

namespace TenderingExpert.Data
{
    public class TenderForm
    {
        public List<SheetInformation> Sheets { get; set; } = new List<SheetInformation>();

        public TenderingInformation TenderInfo { get; set; }

        public PackageInformation PackageInfo { get; set; }

        public string WriteSpace(int num)
        {
            string res = "";
            for (int i = 0; i < num - 1; i++)
                res += " ";

            return res;
        }

        public void Init()
        {
            XmlOperation xml = new XmlOperation();
            Sheets = xml.LoadConfig();
        }

        public void FillContent(ExcelWriter writer)
        {
            foreach (SheetInformation sheetInformation in Sheets)
            {
                var sheetName = sheetInformation.Name;
                writer.AddSheets(sheetName);
                //合并单元格
                foreach (SheetRange range in sheetInformation.MergeCellsPosition)
                    writer.UnitCells(sheetName, range.Left.X, range.Left.Y, range.Right.X, range.Right.Y);
                //填充内容
                foreach (KeyValuePair<SheetCellPosition, CellProperty> pair in sheetInformation.Information)
                {
                    writer.SetCellValue(sheetName, pair.Key.X, pair.Key.Y, pair.Value.Value);
                    writer.SetCellProperty(sheetName, pair.Key.X, pair.Key.Y, pair.Value.FontSize, pair.Value.FontName,
                        pair.Value.FontAlignment, pair.Value.Bold);
                }
                //边框
                foreach (SheetRange range in sheetInformation.BorderArea)
                    writer.SetCellBorder(sheetName, range.Left.X, range.Left.Y, range.Right.X, range.Right.Y);
            }
        }
    }

    public class SheetInformation
    {
        public string Name { get; set; }
        public List<SheetRange> MergeCellsPosition { get; set; }
        public List<SheetRange> BorderArea { get; set; }
        public Dictionary<SheetCellPosition, CellProperty> Information { get; set; }
    }

    public class SheetCellPosition
    {
        public int X { get; set; }
        public int Y { get; set; }
    }

    public class SheetRange
    {
        public SheetCellPosition Left { get; set; }
        public SheetCellPosition Right { get; set; }
    }

    public class CellProperty
    {
        public string Value { get; set; }
        public string FontName { get; set; }
        public int FontSize { get; set; }
        public Alignment FontAlignment { get; set; }
        public bool Bold { get; set; }
    }
}
