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
            InitTemplate();
            InitDataSheet1();
            InitDataSheet2();
            InitDataSheet3();
            InitDataSheet5();
            InitDataSheet6();
            InitDataSheet7();
        }

        private void InitTemplate()
        {
            XmlOperation xml = new XmlOperation();
            var date = TenderInfo.TenderingDate.Substring(0, TenderInfo.TenderingDate.IndexOf('日') + 1);
            Sheets = xml.LoadConfig(PackageInfo.PurchaseInformations.Count, date);
        }

        private void InitDataSheet1()
        {
            var sheet = Sheets.Find(information => information.Name.Contains("表1"));
            sheet.Information.Add(new SheetCellPosition { X = 2, Y = 1 },
                new CellProperty
                {
                    Value = $"项目名称：{TenderInfo.ProjectName}" + WriteSpace(55) + $"项目编号：{TenderInfo.ProjectCode}"
                });

            sheet.Information.Add(new SheetCellPosition { X = 3, Y = 1 },
                new CellProperty
                {
                    Value = $"开标地点：{TenderInfo.TenderingAddress}" + WriteSpace(55) + $"开标时间：{TenderInfo.TenderingDate}"
                });

            sheet.Information.Add(new SheetCellPosition { X = 4, Y = 1 },
                new CellProperty
                {
                    Value = PackageInfo.DeviceName
                });

            var purchaseCount = PackageInfo.PurchaseInformations.Count;
            for (int i = 0; i < purchaseCount; i++)
            {
                var purchase = PackageInfo.PurchaseInformations[i];
                //投标人
                sheet.Information.Add(new SheetCellPosition { X = 6 + i, Y = 2 },
                    new CellProperty
                    {
                        Value = purchase.CompanyName,
                        FontSize = 14
                    });
            }
        }

        private void InitDataSheet2()
        {
            var sheet = Sheets.Find(information => information.Name.Contains("表2"));
            sheet.Information.Add(new SheetCellPosition { X = 2, Y = 1 },
                new CellProperty
                {
                    Value = $"项目名称：{TenderInfo.ProjectName}" + WriteSpace(55) + $"项目编号：{TenderInfo.ProjectCode}"
                });

            sheet.Information.Add(new SheetCellPosition { X = 3, Y = 1 },
                new CellProperty
                {
                    Value = PackageInfo.DeviceName
                });

            var purchaseCount = PackageInfo.PurchaseInformations.Count;
            for (int i = 0; i < purchaseCount; i++)
            {
                var purchase = PackageInfo.PurchaseInformations[i];
                //投标人
                sheet.Information.Add(new SheetCellPosition {X = 4, Y = 3 + i},
                    new CellProperty
                    {
                        Value = purchase.CompanyName,
                        FontSize = 11,
                        FontAlignment = Alignment.Middle
                    });
            }
        }

        private void InitDataSheet3()
        {
            var sheet = Sheets.Find(information => information.Name.Contains("表3"));
            sheet.Information.Add(new SheetCellPosition { X = 2, Y = 1 },
                new CellProperty
                {
                    Value = $"项目名称：{TenderInfo.ProjectName}" + WriteSpace(55) + $"项目编号：{TenderInfo.ProjectCode}"
                });

            sheet.Information.Add(new SheetCellPosition { X = 3, Y = 1 },
                new CellProperty
                {
                    Value = PackageInfo.DeviceName
                });

            var purchaseCount = PackageInfo.PurchaseInformations.Count;
            for (int i = 0; i < purchaseCount; i++)
            {
                var purchase = PackageInfo.PurchaseInformations[i];
                //投标人
                sheet.Information.Add(new SheetCellPosition {X = 5 + i, Y = 1},
                    new CellProperty
                    {
                        Value = purchase.CompanyName,
                        FontSize = 12
                    });
            }
        }
        private void InitDataSheet5()
        {
            var sheet = Sheets.Find(information => information.Name.Contains("表5"));
            sheet.Information.Add(new SheetCellPosition {X = 2, Y = 1},
                new CellProperty($"项目名称：{TenderInfo.ProjectName}" + WriteSpace(20) + $"项目编号：{TenderInfo.ProjectCode}",
                    11));

            sheet.Information.Add(new SheetCellPosition {X = 3, Y = 1},
                new CellProperty(PackageInfo.DeviceName, 11));

            var purchaseCount = PackageInfo.PurchaseInformations.Count;
            for (int i = 0; i < purchaseCount; i++)
            {
                var purchase = PackageInfo.PurchaseInformations[i];
                //投标人
                sheet.Information.Add(new SheetCellPosition { X = 4, Y = 3 + i },
                    new CellProperty
                    {
                        Value = purchase.CompanyName,
                        FontSize = 11,
                        FontAlignment = Alignment.Middle
                    });
            }
        }

        private void InitDataSheet6()
        {
            var sheet = Sheets.Find(information => information.Name.Contains("表6"));
            sheet.Information.Add(new SheetCellPosition {X = 2, Y = 1},
                new CellProperty($"项目名称：{TenderInfo.ProjectName}" + WriteSpace(20) + $"项目编号：{TenderInfo.ProjectCode}"));

            sheet.Information.Add(new SheetCellPosition { X = 3, Y = 1 },
                new CellProperty(PackageInfo.DeviceName));

            var purchaseCount = PackageInfo.PurchaseInformations.Count;
            for (int i = 0; i < purchaseCount; i++)
            {
                var purchase = PackageInfo.PurchaseInformations[i];
                //投标人
                sheet.Information.Add(new SheetCellPosition { X = 5 + i, Y = 1 },
                    new CellProperty
                    {
                        Value = purchase.CompanyName,
                        FontSize = 11,
                        FontAlignment = Alignment.Middle
                    });
            }
        }

        private void InitDataSheet7()
        {
            var sheet = Sheets.Find(information => information.Name.Contains("表7"));
            sheet.Information.Add(new SheetCellPosition {X = 2, Y = 1},
                new CellProperty($"项目名称：{TenderInfo.ProjectName}" + WriteSpace(20) + $"项目编号：{TenderInfo.ProjectCode}"));

            sheet.Information.Add(new SheetCellPosition { X = 3, Y = 1 },
                new CellProperty(PackageInfo.DeviceName));

            var name = PackageInfo.DeviceName.Split('：').Length > 1
                ? PackageInfo.DeviceName.Split('：')[1]
                : PackageInfo.DeviceName;
            sheet.Information.Add(new SheetCellPosition(6, 2),
                new CellProperty($"{name}，{PackageInfo.Quantity}，型号：，单价：人民币 元", 12));
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
                //边框
                foreach (SheetRange range in sheetInformation.BorderArea)
                    writer.SetCellBorder(sheetName, range.Left.X, range.Left.Y, range.Right.X, range.Right.Y);
                //填充内容
                foreach (KeyValuePair<SheetCellPosition, CellProperty> pair in sheetInformation.Information)
                {
                    writer.SetCellValue(sheetName, pair.Key.X, pair.Key.Y, pair.Value.Value);
                    writer.SetCellProperty(sheetName, pair.Key.X, pair.Key.Y, pair.Value.FontSize, pair.Value.FontName,
                        pair.Value.FontAlignment, pair.Value.Bold);
                }
            }
        }
    }

    public class SheetInformation
    {
        public string Name { get; set; }
        public List<SheetRange> MergeCellsPosition { get; set; } = new List<SheetRange>();
        public List<SheetRange> BorderArea { get; set; } = new List<SheetRange>();
        public Dictionary<SheetCellPosition, CellProperty> Information { get; set; } = new Dictionary<SheetCellPosition, CellProperty>();
    }

    public class SheetCellPosition
    {
        public int X { get; set; }
        public int Y { get; set; }

        public SheetCellPosition() { }
        public SheetCellPosition(int x, int y)
        {
            X = x;
            Y = y;
        }
    }

    public class SheetRange
    {
        public SheetCellPosition Left { get; set; }
        public SheetCellPosition Right { get; set; }

        public SheetRange() { }
        public SheetRange(int leftX, int leftY, int rightX, int rightY)
        {
            Left = new SheetCellPosition(leftX, leftY);
            Right = new SheetCellPosition(rightX, rightY);
        }
    }

    public class CellProperty
    {
        public string Value { get; set; }
        public string FontName { get; set; } = "宋体";
        public int FontSize { get; set; } = 10;
        public Alignment FontAlignment { get; set; }
        public bool Bold { get; set; }

        public CellProperty() { }

        public CellProperty(string value, int fontSize = 10, Alignment fontAlignment = Alignment.Left,
            bool bold = false, string fontName = "宋体")
        {
            Value = value;
            FontAlignment = fontAlignment;
            FontName = fontName;
            FontSize = fontSize;
            Bold = bold;
        }
    }
}
