using OneNoteAddin.Setting;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Forms = System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace OneNoteAddin.Handler
{
    /// <summary>
    /// word 的设置类, 需要安装Office Word 2016
    /// </summary>
    public class WordHandler
    {
        private Application wordApplication;
        private Document wordDocument;
        private Range wordRange;

        /// <summary>
        /// 粘贴然后复制
        /// </summary>
        public void PasteAndCopy()
        {
            StartWord();
            wordRange.PasteAndFormat(WdRecoveryType.wdPasteDefault);
            wordRange.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordRange.Font.Shading.Texture = WdTextureIndex.wdTextureNone;

            object unit = WdUnits.wdLine;
            object what = WdGoToItem.wdGoToLine;
            object which = WdGoToDirection.wdGoToLast;
            object count = 99999999;
            wordApplication.Selection.GoTo(ref what, ref which, ref count);
            RemoveTailSpace();

            // 不复制最后的换行
            var sentences = wordDocument.Sentences;
            wordRange.SetRange(sentences[1].Start, sentences[sentences.Count].End - 1);
            if (sentences.Count != 1 || sentences[1].End != 1)
            {
                // 避免剪切空行的异常
                wordRange.Cut();
            }
        }

        private void StartWord()
        {
            bool shouldStart = wordApplication == null;
            if (!shouldStart)
            {
                // 防止word 进程死掉
                try
                {
                    wordRange.Select();
                }
                catch (Exception err)
                {
                    shouldStart = true;
                }
            }
            if (shouldStart)
            {
                try
                {
                    wordApplication = new Application();
                    wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    wordDocument = wordApplication.Documents.Add();
                    wordRange = wordDocument.Paragraphs[1].Range;
                }
                catch (Exception err)
                {
                    Forms.MessageBox.Show("Error happened while start word : \n" + err.Message);
                }
            }
        }

        private void RemoveTailSpace()
        {
            int count = 0, maxTry = 1000;
            while (count++ < maxTry && Regex.IsMatch(wordApplication.Selection.Text, @"\s|\r?\n"))
            {
                wordApplication.Selection.Delete();
                wordApplication.Selection.MoveLeft();
            }
        }

        /// <summary>
        /// 粘贴代码然后复制到剪切板.
        /// 会给代码添加一个无边框的1*1 的表格, 底纹是 #f8f8f8, 同时会删除最前面的重复的空格.
        /// </summary>
        public void CopyCode()
        {
            StartWord();

            //把代码粘贴在一个表格中
            Table table = wordRange.Tables.Add(wordRange, 1, 1);
            table.Cell(0, 0).Range.PasteAndFormat(WdRecoveryType.wdPasteDefault);
            //table.Cell(0, 0).Range.set_Style(WdBuiltinStyle.wdStyleNormal);

            // 删除末尾的空白
            object unit = WdUnits.wdLine;
            object what = WdGoToItem.wdGoToLine;
            object which = WdGoToDirection.wdGoToLast;
            object count = 99999999;
            wordApplication.Selection.GoTo(ref what, ref which, ref count);
            wordApplication.Selection.MoveLeft();
            wordApplication.Selection.Delete(); // 删除多余的换行
            wordApplication.Selection.MoveLeft();
            wordApplication.Selection.MoveLeft(); // 进入单元格内
            RemoveTailSpace();

            // 删除每行前面的重复空格
            string text = table.Cell(0, 0).Range.Text;
            int spaceCount = int.MaxValue;
            foreach (var line in text.Split('\n'))
            {
                int thisRowCount = 0;
                foreach (var c in line)
                {
                    if (Regex.IsMatch(c.ToString(), @"\s"))
                    {
                        thisRowCount++;
                    }
                    else
                    {
                        break;
                    }
                }
                spaceCount = Math.Min(spaceCount, thisRowCount);
            }
            if (spaceCount > 0)
            {
                table.Cell(0, 0).Range.Select();
                wordApplication.Selection.MoveLeft();
                do
                {
                    for (int i = 0; i < spaceCount; i++)
                    {
                        string s = wordApplication.Selection.Text;
                        if (!@Regex.IsMatch(s, @"\r|\n"))
                        {
                            wordApplication.Selection.Delete();
                        }
                    }
                } while (wordApplication.Selection.MoveDown() != 0);
            }


            // 设置表格底纹颜色
            table.Range.Shading.BackgroundPatternColor = (WdColor)0xf8f8f8;
            //wordDocument.SaveAs2("d:\\test.docx");

            //table.Range.Cut();
            var sentences = wordDocument.Sentences;
            wordRange.SetRange(sentences[1].Start, sentences[sentences.Count].End);
            wordRange.Cut();
        }

        public void Close()
        {
            if (wordApplication != null)
            {
                try
                {
                    wordDocument.Close(null, null, null);
                    wordApplication.Quit();
                    wordDocument = null;
                    wordApplication = null;
                    wordRange = null;
                }
                catch (Exception err)
                {
                    //do nothing
                }
            }
        }

        /// <summary>
        /// 新建一个表格, 然后复制到剪切板中
        /// </summary>
        /// <param name="tableSetting">想要的表格样式</param>
        public void CreateTable(TableModel tableSetting)
        {
            StartWord();

            try
            {
                int row = tableSetting.Row;
                int column = tableSetting.Column;
                if (tableSetting.HeadInLeft)
                {
                    column++;
                }
                else
                {
                    row++;
                }

                Table table = wordRange.Tables.Add(wordRange, row, column);
                table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
                table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                List<Cell> titleCells = new List<Cell>();
                if (tableSetting.HeadInLeft)
                {
                    for (int i = 1; i <= row; i++)
                    {
                        titleCells.Add(table.Cell(i, 1));
                    }
                }
                else
                {
                    for (int i = 1; i <= column; i++)
                    {
                        titleCells.Add(table.Cell(1, i));
                    }
                }
                foreach (var titleCell in titleCells)
                {
                    var range = titleCell.Range;
                    range.Font.Size = 11;
                    range.Font.Bold = 1;
                    range.Font.Color = (WdColor)Convert.ToInt32(tableSetting.ForeColor, 16);
                    range.Shading.BackgroundPatternColor = (WdColor)Convert.ToInt32(tableSetting.BackColor, 16);
                }

                // 剪切到剪切板上
                table.Range.Cut();
                wordRange.Delete();
            }
            catch (Exception err)
            {
                Forms.MessageBox.Show("Error while insert table : \n" + err.ToString());
            }
        }
    }
}
