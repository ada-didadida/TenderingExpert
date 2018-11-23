using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ExcelOperator;

namespace TenderingExpert.Data
{
    public class XmlOperation
    {
        private XmlDocument xmlDocument;

        public XmlOperation()
        {
            xmlDocument = new XmlDocument();
        }

        public List<SheetInformation> LoadConfig()
        {
            var result = new List<SheetInformation>();

            var path = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config"), "ExcelTemplate.xml");
            xmlDocument.Load(path);

            var root = xmlDocument.SelectNodes("/Excel");
            if (root != null)
            {
                foreach (XmlElement element in root)
                {
                    var sheetInformation = new SheetInformation
                    {
                        Name = element.GetElementsByTagName("Name")[0].InnerText,
                        MergeCellsPosition = new List<SheetRange>(),
                        BorderArea = new List<SheetRange>(),
                        Information = new Dictionary<SheetCellPosition, CellProperty>()
                    };

                    var mergeCellsPosition = element.GetElementsByTagName("MergeCellsPosition")[0];
                    foreach (XmlElement pair in mergeCellsPosition)
                    {
                        var left = pair.GetElementsByTagName("Left")[0].Attributes;
                        var right = pair.GetElementsByTagName("Right")[0].Attributes;

                        var range = new SheetRange
                        {
                            Left = new SheetCellPosition{X = Convert.ToInt32(left["X"].Value), Y = Convert.ToInt32(left["Y"].Value)},
                            Right = new SheetCellPosition{ X = Convert.ToInt32(right["X"].Value), Y = Convert.ToInt32(right["Y"].Value)}
                        };

                        sheetInformation.MergeCellsPosition.Add(range);
                    }

                    var borderArea = element.GetElementsByTagName("BorderArea")[0];
                    foreach (XmlElement pair in borderArea)
                    {
                        var left = pair.GetElementsByTagName("Left")[0].Attributes;
                        var right = pair.GetElementsByTagName("Right")[0].Attributes;

                        var range = new SheetRange
                        {
                            Left = new SheetCellPosition { X = Convert.ToInt32(left["X"].Value), Y = Convert.ToInt32(left["Y"].Value) },
                            Right = new SheetCellPosition { X = Convert.ToInt32(right["X"].Value), Y = Convert.ToInt32(right["Y"].Value) }
                        };

                        sheetInformation.BorderArea.Add(range);
                    }

                    var information = element.GetElementsByTagName("Information")[0];
                    foreach (XmlElement cell in information)
                    {
                        var pos = new SheetCellPosition
                        {
                            X = Convert.ToInt32(cell.Attributes["X"].Value),
                            Y = Convert.ToInt32(cell.Attributes["Y"].Value)
                        };

                        var property = new CellProperty
                        {
                            Value = cell.Attributes["Value"].Value,
                            FontName = cell.Attributes["FontName"].Value,
                            FontSize = Convert.ToInt32(cell.Attributes["FontSize"].Value),
                            FontAlignment = (Alignment) Convert.ToInt32(cell.Attributes["FontAlignment"].Value),
                            Bold = Convert.ToBoolean(Convert.ToInt32(cell.Attributes["Bold"].Value))
                        };

                        sheetInformation.Information.Add(pos, property);
                    }

                    result.Add(sheetInformation);
                }
            }

            return result;
        }
    }
}
