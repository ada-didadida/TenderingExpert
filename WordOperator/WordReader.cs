using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace WordOperator
{
    public class WordReader
    {
        private readonly string wordPath;

        private Application wordApplication;

        private Document wordDocument;

        private object unknown = Type.Missing;

        public WordReader(string wordPath)
        {
            var currentPath = Path.Combine(wordPath);
            if (File.Exists(currentPath))
            {
                this.wordPath = wordPath;
                wordApplication = new Application();
            }
            else
                throw new ArgumentException("文件不存在，请重新选择", nameof(wordPath));
        }

        public void StartRead()
        {
            try
            {
                object file = wordPath;

                object readOnly = true;

                wordApplication.Visible = false;

                wordDocument = wordApplication.Documents.Open(ref file, ref unknown, ref readOnly, ref unknown,
                    ref unknown, ref unknown, ref unknown, ref unknown, ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown, ref unknown, ref unknown);
            }
            catch (Exception e)
            {
                throw new Exception("无法打开文档", e);
            }
        }

        public string GetParagraphsContent(int paragraphIndex)
        {
            return wordDocument != null ? wordDocument.Paragraphs[paragraphIndex].Range.Text : string.Empty;
        }

        public int GetParagraphsCount()
        {
            if(wordDocument != null)
                return wordDocument.Paragraphs.Count;

            return 0;
        }

        public string GetAllContent()
        {
            var result = string.Empty;

            foreach (Paragraph paragraph in wordDocument.Paragraphs)
                result += paragraph.Range.Text + "\n";

            return result;
        }

        public string FindKeyValue(string key)
        {
            Range range = wordDocument.Range(0, 0);
            while (range.Find.Execute(key))
            {
                range.MoveEndUntil(Environment.NewLine);
                var result = range.Text.Replace(key, "");
                if(string.IsNullOrEmpty(result))
                    continue;

                return result;
            }

            return null;
        }

        /// <summary>
        /// 获取表格内容
        /// 所有索引从1开始！
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public List<List<string>> GetTableContent(int index)
        {
            if (wordDocument != null)
            {
                var result = new List<List<string>>();

                var table = wordDocument.Tables[index];
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    var row = new List<string>();
                    for (int j = 1; j <= table.Columns.Count; j++)
                        row.Add(table.Cell(i, j).Range.Text);

                    result.Add(row);
                }

                return result;
            }

            return null;
        }

        public void Close()
        {
            if (wordApplication == null) return;

            wordDocument?.Close(ref unknown, ref unknown, ref unknown);
            wordApplication.Quit(ref unknown, ref unknown, ref unknown);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApplication);
            wordDocument = null;
            wordApplication = null;

            GC.Collect();
        }
    }
}
