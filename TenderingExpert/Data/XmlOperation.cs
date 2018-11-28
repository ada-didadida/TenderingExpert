using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using ExcelOperator;

namespace TenderingExpert.Data
{
    public class XmlOperation
    {
        private XmlDocument xmlDocument;

        private bool isOnePackage;

        public XmlOperation()
        {
            xmlDocument = new XmlDocument();
        }

        public void SetPackageFlag(bool onePackage)
        {
            isOnePackage = onePackage;
        }

        public List<SheetInformation> LoadConfig(int purchaseCount, string date)
        {
            OpenXmlConfig();
            var res = new List<SheetInformation>
            {
                LoadSheet1(purchaseCount),
                LoadSheet2(purchaseCount, date),
                LoadSheet3(purchaseCount, date),
                LoadSheet5(purchaseCount, date),
                LoadSheet6(purchaseCount, date),
                LoadSheet7(date)
            };

            xmlDocument = null;

            return res;
        }

        public void OpenXmlConfig(string file = "Config/ExcelTemplate.xml")
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, file);
            if (File.Exists(path))
            {
                try
                {
                    xmlDocument.Load(path);
                }
                catch (Exception e)
                {
                    throw new XmlException($"Can not load config file. {e.Message}", e);
                }
            }
        }

        private XmlElement GetSheetElement(string sheetName)
        {
            var root = xmlDocument.DocumentElement;
            if (root != null)
            {
                foreach (XmlElement element in root)
                {
                    var name = element.Attributes["Name"]?.Value;
                    if (name != null && (name == sheetName || name.Contains(sheetName)))
                        return element;
                }
            }

            return null;
        }

        private string GetAttribute(XmlElement element, string attribute)
        {
            if (element.HasAttribute(attribute))
            {
                return element.GetAttribute(attribute);
            }

            return string.Empty;
        }

        private SheetInformation LoadSheet1(int purchaseCount)
        {
            var info = new SheetInformation();

            var sheet = GetSheetElement("表1");
            if (sheet != null)
            {
                int width = 6;
                int mergeRowCount = 3;
                SetCommonProperty(sheet, info, mergeRowCount, width);

                //空一行开始表格
                int formStartRow = mergeRowCount + 2;
                if (isOnePackage)
                {
                    formStartRow = mergeRowCount + 1;
                }
                //表头
                var formHeader = sheet.GetElementsByTagName("Header")[0];
                for (int i = 0; i < formHeader.ChildNodes.Count; i++)
                {
                    var header = formHeader.ChildNodes[i].InnerText;
                    info.Information.Add(new SheetCellPosition(formStartRow, i + 1),
                        new CellProperty(header, 14, Alignment.Middle));
                }

                for (int i = 0; i < purchaseCount; i++)
                {
                    //编号
                    info.Information.Add(new SheetCellPosition(formStartRow + i + 1, 1),
                        new CellProperty((i + 1).ToString(), 14, Alignment.Middle));
                    //内容
                    info.Information.Add(new SheetCellPosition
                    {
                        X = formStartRow + i + 1,
                        Y = 4
                    }, new CellProperty
                    {
                        Value = "有",
                        FontAlignment = Alignment.Middle,
                        FontSize = 14
                    });
                    info.Information.Add(new SheetCellPosition
                    {
                        X = formStartRow + i + 1,
                        Y = 5
                    }, new CellProperty
                    {
                        Value = "无",
                        FontAlignment = Alignment.Middle,
                        FontSize = 14
                    });
                }
                //结尾位置
                int endRow = formStartRow + purchaseCount + 2;

                info.MergeCellsPosition.Add(new SheetRange
                {
                    Left = new SheetCellPosition {X = endRow, Y = 1},
                    Right = new SheetCellPosition {X = endRow, Y = width}
                });

                var endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[0].InnerText;
                info.Information.Add(new SheetCellPosition
                {
                    X = endRow, Y = 1
                }, new CellProperty
                {
                    Value = endContent,
                    FontSize = 12
                });
                //边框
                info.BorderArea.Add(new SheetRange
                {
                    Left = new SheetCellPosition { X = formStartRow, Y = 1},
                    Right = new SheetCellPosition { X = formStartRow + purchaseCount, Y = width}
                });
            }
            else
            {
                throw new NullReferenceException("Can not Find Sheet 1");
            }

            return info;
        }

        private SheetInformation LoadSheet2(int purchaseCount, string date)
        {
            var info = new SheetInformation();

            var sheet = GetSheetElement("表2");
            if (sheet != null)
            {
                int width = 2 + purchaseCount;
                int mergeRowCount = 2;
                SetCommonProperty(sheet, info, mergeRowCount, width);

                //空一行开始表格
                int formStartRow = mergeRowCount + 2;
                if (isOnePackage)
                    formStartRow = mergeRowCount + 1;
                //表头
                var formHeader = sheet.GetElementsByTagName("Header")[0];
                for (int i = 0; i < formHeader.ChildNodes.Count; i++)
                {
                    var header = formHeader.ChildNodes[i].InnerText;
                    info.Information.Add(new SheetCellPosition
                    {
                        X = formStartRow,
                        Y = i + 1
                    }, new CellProperty
                    {
                        Value = header,
                        FontAlignment = Alignment.Middle,
                        FontSize = 11,
                        Bold = i == 0
                    });
                }
                //内容
                int contentCount = 0;
                var formContent = sheet["FormInformation"]?["Content"];
                if (formContent != null)
                {
                    contentCount = formContent.ChildNodes.Count;
                    for (int i = 0; i < contentCount; i++)
                    {
                        var value = formContent.ChildNodes[i].InnerText;
                        //编号
                        info.Information.Add(new SheetCellPosition
                        {
                            X = formStartRow + i + 1, Y = 1
                        }, new CellProperty
                        {
                            Value = (i + 1).ToString(),
                            FontAlignment = Alignment.Middle,
                        });
                        //内容
                        info.Information.Add(new SheetCellPosition {X = formStartRow + i + 1, Y = 2},
                            new CellProperty {Value = value});
                        //对勾
                        for (int j = 0; j < purchaseCount; j++)
                        {
                            info.Information.Add(new SheetCellPosition { X = formStartRow + i + 1, Y = j + 3 },
                                new CellProperty { Value = "√", FontAlignment = Alignment.Middle});
                        }
                    }
                }
                //结论
                info.Information.Add(new SheetCellPosition { X = formStartRow + contentCount + 1, Y = 2 },
                    new CellProperty { Value = "结论", Bold = true });
                for (int i = 0; i < purchaseCount; i++)
                {
                    info.Information.Add(new SheetCellPosition { X = formStartRow + contentCount + 1, Y = i + 3 },
                        new CellProperty { Value = "合格", FontAlignment = Alignment.Middle });
                }
                //结尾位置
                int endRow = formStartRow + contentCount + 2;

                var endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[0].InnerText;
                info.Information.Add(new SheetCellPosition {X = endRow, Y = 1},
                    new CellProperty {Value = endContent});

                endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[1].InnerText;
                info.Information.Add(new SheetCellPosition { X = endRow +1, Y = 1 },
                    new CellProperty { Value = endContent, FontSize = 12 });
                //Date
                info.Information.Add(new SheetCellPosition { X = endRow + 1, Y = width },
                    new CellProperty { Value = $"日期：{date}", FontSize = 12, FontAlignment = Alignment.Right});

                //边框
                info.BorderArea.Add(new SheetRange
                {
                    Left = new SheetCellPosition { X = formStartRow, Y = 1},
                    Right = new SheetCellPosition { X = formStartRow + contentCount + 1, Y = width }
                });
            }
            else
            {
                throw new NullReferenceException("Can not Find Sheet 2");
            }

            return info;
        }

        private SheetInformation LoadSheet3(int purchaseCount, string date)
        {
            var info = new SheetInformation();

            var sheet = GetSheetElement("表3");
            if (sheet != null)
            {
                int width = 9;
                int mergeRowCount = 2;
                SetCommonProperty(sheet, info, mergeRowCount, width);

                //空一行开始表格
                int formStartRow = mergeRowCount + 2;
                if (isOnePackage)
                    formStartRow = mergeRowCount + 1;
                //表头
                var formHeader = sheet.GetElementsByTagName("Header")[0];
                for (int i = 0; i < formHeader.ChildNodes.Count; i++)
                {
                    var header = formHeader.ChildNodes[i].InnerText;
                    info.Information.Add(new SheetCellPosition
                    {
                        X = formStartRow,
                        Y = i + 1
                    }, new CellProperty
                    {
                        Value = header,
                        FontAlignment = Alignment.Middle,
                        FontSize = 12,
                        Bold = i != 0
                    });
                }
                //对勾
                for (int i = 0; i < purchaseCount; i++)
                {
                    for (int j = 2; j < 9; j++)
                    {
                        info.Information.Add(new SheetCellPosition { X = formStartRow + i + 1, Y = j },
                            new CellProperty { Value = "√", FontSize = 12 });
                    }
                    //结论
                    info.Information.Add(new SheetCellPosition { X = formStartRow + i + 1, Y = 9 },
                        new CellProperty { Value = "合格", FontSize = 12 });
                }
                //结尾位置
                int endRow = formStartRow + purchaseCount + 2;

                var endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[0].InnerText;
                info.Information.Add(new SheetCellPosition { X = endRow, Y = 1 },
                    new CellProperty { Value = endContent, FontSize = 12 });
                //日期
                info.MergeCellsPosition.Add(new SheetRange
                {
                    Left = new SheetCellPosition {X = endRow, Y = width - 2},
                    Right = new SheetCellPosition {X = endRow, Y = width}
                });
                info.Information.Add(new SheetCellPosition {X = endRow, Y = width - 2},
                    new CellProperty {Value = $"日期：{date}", FontSize = 12, FontAlignment = Alignment.Middle});
                //边框
                info.BorderArea.Add(new SheetRange
                {
                    Left = new SheetCellPosition { X = formStartRow, Y = 1 },
                    Right = new SheetCellPosition { X = formStartRow + purchaseCount, Y = width }
                });
            }
            else
            {
                throw new NullReferenceException("Can not Find Sheet 3");
            }

            return info;
        }

        private SheetInformation LoadSheet5(int purchaseCount, string date)
        {
            var info = new SheetInformation();

            var sheet = GetSheetElement("表5");
            if (sheet != null)
            {
                int width = 2 + purchaseCount;
                int mergeRowCount = 2;
                SetCommonProperty(sheet, info, mergeRowCount, width);
                //合并
                if (!isOnePackage)
                    info.MergeCellsPosition.Add(new SheetRange(mergeRowCount + 1, 1, mergeRowCount + 1, 3));
                //空一行开始表格
                int formStartRow = mergeRowCount + 2;
                if (isOnePackage)
                    formStartRow = mergeRowCount + 1;
                //表头
                var formHeader = sheet.GetElementsByTagName("Header")[0];
                for (int i = 0; i < formHeader.ChildNodes.Count; i++)
                {
                    var header = formHeader.ChildNodes[i].InnerText;
                    info.Information.Add(new SheetCellPosition(formStartRow, i + 1),
                        new CellProperty(header, 12, Alignment.Middle, true));
                }
                //内容
                int judgeCount = Convert.ToInt32(sheet.GetElementsByTagName("JudgesCount")[0].InnerText);
                for (int i = 0; i < judgeCount; i++)
                {
                    info.Information.Add(new SheetCellPosition(formStartRow + i + 1, 1),
                        new CellProperty((i + 1).ToString(), 12, Alignment.Middle));
                }

                info.Information.Add(new SheetCellPosition(formStartRow + judgeCount + 1, 2),
                    new CellProperty("平均得分", 12, Alignment.Middle, true));
                //公式
                for (int i = 0; i < purchaseCount; i++)
                {
                    char column = (char) ('C' + i);
                    info.Information.Add(new SheetCellPosition(formStartRow + judgeCount + 1, 3 + i),
                        new CellProperty($"=AVERAGE({column.ToString()}5:{column.ToString()}{formStartRow + judgeCount})", 12,
                            Alignment.Middle, true));
                }

                //结尾位置
                int endRow = formStartRow + judgeCount + 2;

                info.MergeCellsPosition.Add(new SheetRange(endRow, 1, endRow, 3));

                var endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[0].InnerText;
                info.Information.Add(new SheetCellPosition { X = endRow, Y = 1 },
                    new CellProperty { Value = endContent, FontSize = 11 });
                //日期
                info.Information.Add(new SheetCellPosition(endRow, width),
                    new CellProperty($"日期：{date}", 11, Alignment.Right));
                //边框
                info.BorderArea.Add(new SheetRange
                {
                    Left = new SheetCellPosition {X = formStartRow, Y = 1},
                    Right = new SheetCellPosition {X = formStartRow + judgeCount + 1, Y = width}
                });
            }
            else
            {
                throw new NullReferenceException("Can not Find Sheet 5");
            }

            return info;
        }
        private SheetInformation LoadSheet6(int purchaseCount, string date)
        {
            var info = new SheetInformation();

            var sheet = GetSheetElement("表6");
            if (sheet != null)
            {
                int judgeCount = Convert.ToInt32(sheet.GetElementsByTagName("JudgesCount")[0].InnerText);
                int width = 3 + judgeCount;
                int mergeRowCount = 2;
                SetCommonProperty(sheet, info, mergeRowCount, width);
                //空一行开始表格
                int formStartRow = mergeRowCount + 2;
                if (isOnePackage)
                    formStartRow = mergeRowCount + 1;
                //表头
                var formHeader = sheet.GetElementsByTagName("Header")[0];
                for (int i = 0; i < formHeader.ChildNodes.Count; i++)
                {
                    var header = formHeader.ChildNodes[i].InnerText;
                    info.Information.Add(new SheetCellPosition(formStartRow, i + 1),
                        new CellProperty(header, 14, Alignment.Middle, true));
                }
                //内容
                info.Information.Add(new SheetCellPosition(formStartRow, width - 1),
                    new CellProperty("平均得分", 14, Alignment.Middle, true));
                info.Information.Add(new SheetCellPosition(formStartRow, width),
                    new CellProperty("排序", 12, Alignment.Middle));
                //公式
                for (int i = 0; i < purchaseCount; i++)
                {
                    info.Information.Add(new SheetCellPosition(formStartRow + i + 1, width - 1),
                        new CellProperty(
                            $"=AVERAGE({'B'.ToString()}{formStartRow + i + 1}:{((char) ('B' + judgeCount - 1)).ToString()}{formStartRow + i + 1})",
                            12,
                            Alignment.Middle, true));
                }

                //结尾位置
                int endRow = formStartRow + purchaseCount + 1;                

                var endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[0].InnerText;
                info.Information.Add(new SheetCellPosition { X = endRow, Y = 1 },
                    new CellProperty { Value = endContent, FontSize = 12 });
                //日期
                info.Information.Add(new SheetCellPosition(endRow, width),
                    new CellProperty($"日期：{date}", 12, Alignment.Right));
                //边框
                info.BorderArea.Add(new SheetRange
                {
                    Left = new SheetCellPosition { X = formStartRow, Y = 1 },
                    Right = new SheetCellPosition { X = formStartRow + purchaseCount, Y = width }
                });
            }
            else
            {
                throw new NullReferenceException("Can not Find Sheet 6");
            }

            return info;
        }

        private SheetInformation LoadSheet7(string date)
        {
            var info = new SheetInformation();

            var sheet = GetSheetElement("表7");
            if (sheet != null)
            {
                int width = 2;
                int mergeRowCount = 2;
                SetCommonProperty(sheet, info, mergeRowCount, width);

                //空一行开始表格
                int formStartRow = mergeRowCount + 2;
                if (isOnePackage)
                    formStartRow = mergeRowCount + 1;
                //表头
                var formHeader = sheet.GetElementsByTagName("Header")[0];
                for (int i = 0; i < formHeader.ChildNodes.Count; i++)
                {
                    var header = formHeader.ChildNodes[i].InnerText;
                    info.Information.Add(new SheetCellPosition(formStartRow + i, 1),
                        new CellProperty(header, 12));
                }
                //内容
                info.Information.Add(new SheetCellPosition(isOnePackage ? 7 : 8, 2),
                    new CellProperty("人名币  元", 12));
                //结尾位置
                int endRow = formStartRow + formHeader.ChildNodes.Count;

                var endContent = sheet.GetElementsByTagName("End")[0].ChildNodes[0].InnerText;
                info.Information.Add(new SheetCellPosition { X = endRow, Y = 1 },
                    new CellProperty { Value = endContent, FontSize = 11, Bold = true});
                //日期
                info.Information.Add(new SheetCellPosition(endRow, width),
                    new CellProperty($"日期：{date}", 11, Alignment.Right));
                //边框
                info.BorderArea.Add(new SheetRange
                {
                    Left = new SheetCellPosition {X = formStartRow, Y = 1},
                    Right = new SheetCellPosition {X = endRow - 1, Y = width}
                });
            }
            else
            {
                throw new NullReferenceException("Can not Find Sheet 7");
            }

            return info;
        }

        private void SetCommonProperty(XmlElement sheet, SheetInformation sheetInfo, int mergeRowCount, int width)
        {
            if (sheet != null)
            {
                sheetInfo.Name = GetAttribute(sheet, "Name");

                //添加初始位置合并单元格
                for (int i = 0; i < mergeRowCount; i++)
                {
                    sheetInfo.MergeCellsPosition.Add(new SheetRange
                    {
                        Left = new SheetCellPosition { X = i + 1, Y = 1 },
                        Right = new SheetCellPosition { X = i + 1, Y = width }
                    });
                }
                //标题
                var title = sheet.GetElementsByTagName("Title")[0].InnerText;
                sheetInfo.Information.Add(new SheetCellPosition { X = 1, Y = 1 }, new CellProperty
                {
                    Value = title,
                    Bold = true,
                    FontAlignment = Alignment.Middle,
                    FontSize = 14
                });
            }
        }
    }
}
