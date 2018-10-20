using OneNoteAddin.Setting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OneNoteAddin.Handler
{
    /// <summary>
    /// word 的设置类, 需要安装Office Word 2016
    /// </summary>
    public class WordHandler
    {
        private Word.Application wordApplication;
        private Word.Document wordDocument;
        private Word.Range wordRange;

        /// <summary>
        /// 粘贴然后复制
        /// </summary>
        public void PasteAndCopy()
        {
            StartWord();
            wordRange.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);
            wordRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            object unit = Word.WdUnits.wdLine;
            object what = Word.WdGoToItem.wdGoToLine;
            object which = Word.WdGoToDirection.wdGoToLast;
            object count = 99999999;
            wordApplication.Selection.GoTo(ref what, ref which, ref count);
            RemoveTailSpace();

            WordCut();
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
                    wordApplication = new Word.Application();
                    wordApplication.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    wordDocument = wordApplication.Documents.Add();
                    wordRange = wordDocument.Paragraphs[1].Range;
                }
                catch (Exception err)
                {
                    MessageBox.Show("Error happened while start word : \n" + err.Message);
                }
            }
        }

        private void RemoveTailSpace()
        {
            int count = 0, maxTry = 1000;
            while (count++ < maxTry && Regex.IsMatch(wordApplication.Selection.Text, "\\s"))
            {
                wordApplication.Selection.Delete();
                wordApplication.Selection.MoveLeft();
            }
        }

        public void WordCut()
        {
            wordRange.Font.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            wordRange.Cut();
        }

        /// <summary>
        /// 粘贴代码然后复制到剪切板.
        /// 会给代码添加一个无边框的1*1 的表格, 底纹是 #f8f8f8, 同时会删除最前面的重复的空格.
        /// </summary>
        public void CopyCode()
        {
            StartWord();

            //把代码粘贴在一个表格中
            Word.Table table = wordRange.Tables.Add(wordRange, 1, 1);
            table.Cell(0, 0).Range.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);
            //table.Cell(0, 0).Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            // 删除末尾的空白
            object unit = Word.WdUnits.wdLine;
            object what = Word.WdGoToItem.wdGoToLine;
            object which = Word.WdGoToDirection.wdGoToLast;
            object count = 99999999;
            wordApplication.Selection.GoTo(ref what, ref which, ref count);
            wordApplication.Selection.MoveLeft();
            wordApplication.Selection.MoveLeft();
            wordApplication.Selection.MoveLeft();
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
            table.Range.Shading.BackgroundPatternColor = (Word.WdColor)0xf8f8f8;
            //wordDocument.SaveAs2("d:\\test.docx");

            WordCut();
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

                Word.Table table = wordRange.Tables.Add(wordRange, row, column);
                table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                List<Word.Cell> titleCells = new List<Word.Cell>();
                if (tableSetting.HeadInLeft)
                {
                    for (int i = 0; i < row; i++)
                    {
                        titleCells.Add(table.Cell(i, 0));
                    }
                }
                else
                {
                    for (int i = 0; i < row; i++)
                    {
                        titleCells.Add(table.Cell(0, i));
                    }
                }
                foreach (var titleCell in titleCells)
                {
                    var range = titleCell.Range;
                    range.Font.Size = 11;
                    range.Font.Bold = 1;
                    range.Font.Color = (Word.WdColor)Convert.ToInt32(tableSetting.ForeColor, 16);
                    range.Shading.BackgroundPatternColor = (Word.WdColor)Convert.ToInt32(tableSetting.BackColor, 16);
                }

                // 剪切到剪切板上
                wordRange.Cut();
            }
            catch (Exception err)
            {
                MessageBox.Show("Error while insert table : \n" + err.ToString());
            }
        }
    }
}
