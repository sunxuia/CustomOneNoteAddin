using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using OneNoteAddin.Handler;
using OneNoteAddin.Setting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;
using Word = Microsoft.Office.Interop.Word;

namespace OneNoteAddin
{
    /// <summary>
    /// ribbon events
    /// </summary>
    public partial class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        WordHandler wordHandler = new WordHandler();

        VSCodeHandler vsCodeHandler;

        private string prevCodeStyle;

        public void OnTextChange(IRibbonControl control, string newText)
        {
            SetDefaultValue(GetControlId(control), newText);
        }

        private void SetDefaultValue(string id, string value)
        {
            foreach (var defaultValue in setting.DefaultValues)
            {
                if (defaultValue.Id == id)
                {
                    defaultValue.Value = value;
                    return;
                }
            }
            setting.DefaultValues.Add(new DefaultValueModel()
            {
                Id = id,
                Value = value
            });
        }

        public void OnPasteWithWord(IRibbonControl control)
        {
            wordHandler.PasteAndCopy();
            SendKeys.SendWait("^(v)");
        }

        public void OnInsertCode(IRibbonControl control)
        {
            FormatByVSCode();
            vsCodeHandler.windowHandler.DeactiveWindow();
            wordHandler.CopyCode();
            // 粘贴并去掉末尾的空格
            SendKeys.SendWait("^(v){BACKSPACE}{BACKSPACE}");

        }

        private void FormatByVSCode()
        {
            string codeStyle = GetDefaultValue("cmbStyle");
            if (!string.IsNullOrEmpty(codeStyle) && codeStyle.ToUpper() != "ORIGINAL FORMAT")
            {
                if (prevCodeStyle == null)
                {
                    vsCodeHandler.InitialForFormat();
                }
                if (prevCodeStyle != codeStyle)
                {
                    vsCodeHandler.ChangeLanguageMode(codeStyle);
                    prevCodeStyle = codeStyle;
                }
                vsCodeHandler.PasteFormatCut();
            }
        }


        public void OnInsertComment(IRibbonControl control)
        {
            string codeStyle = GetDefaultValue("cmbStyle");
            if (!string.IsNullOrEmpty(codeStyle) && codeStyle.ToUpper() != "ORIGINAL FORMAT")
            {
                string comment = "";
                foreach (var style in setting.CodeStyles)
                {
                    if (style.Label == codeStyle)
                    {
                        comment = style.Comment;
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(comment))
                {
                    CopyToClipboard(comment);
                    FormatByVSCode();
                    wordHandler.PasteAndCopy();
                    SendKeys.SendWait("^(v)");
                }
            }
        }

        public void OnOpenInVSCode(IRibbonControl control)
        {
            OneNotePageHandler page = new OneNotePageHandler(app);
            string text = page.GetSelectedText();
            if (string.IsNullOrEmpty(text))
            {
                var table = page.GetCursorElement("Table");
                if (table == null)
                {
                    MessageBox.Show("Please select text or set input cursor into a table.");
                    return;
                }
                else
                {
                    text = page.GetInnerText(table);
                }
            }

            CopyToClipboard(text);
            VSCodeHandler codeHandler = new VSCodeHandler(setting.VSCode);
            if (codeHandler.EditCode(GetDefaultValue("cmbStyle"), out string newText))
            {
                CopyToClipboard(newText);
                FormatByVSCode();
                wordHandler.PasteAndCopy();
                SendKeys.SendWait("^(v)");
            }
        }

        public void OnSetFontClick(IRibbonControl control)
        {
            OneNotePageHandler page = new OneNotePageHandler(app);
            string id = GetControlId(control).ToUpper();
            string fontName = GetDefaultValue("cmbFont" + Regex.Match(id, @"\d+").Value);
            bool addFontFamilyIfNotExist = false;
            Func<string, string> setFontFamily = value =>
            {
                value = value ?? "";
                var match = Regex.Match(value, @"font-family:[^;]*;");
                if (match.Success)
                {
                    value = value.Remove(match.Groups[0].Index, match.Groups[0].Length);
                    value = $"font-family:{fontName};" + value;
                }
                else if (addFontFamilyIfNotExist)
                {
                    value = $"font-family:{fontName};" + value;
                }
                return value;
            };

            Action<HtmlAgilityPack.HtmlDocument> changeTextHtml = doc =>
            {
                foreach (var node in doc.DocumentNode.ChildNodes)
                {
                    foreach (var attr in node.Attributes)
                    {
                        if (attr.Name == "style")
                        {
                            attr.Value = setFontFamily.Invoke(attr.Value);
                            break;
                        }
                    }
                }
            };

            if (id.Contains("SELECTION"))
            {
                // set selection
                var selection = page.GetSelectedElements("OE");
                addFontFamilyIfNotExist = true;
                page.SetAttributeValue(
                    selection,
                    "style",
                    (ref bool exist, ref string value) =>
                    {
                        exist = true;
                        value = setFontFamily(value);
                    });
                addFontFamilyIfNotExist = false;
                page.EnumTextHtml(selection, changeTextHtml);
            }
            else
            {
                // set page
                addFontFamilyIfNotExist = false;
                page.SetPageAttributes("style", (ref bool exist, ref string value) =>
                {
                    if (exist)
                    {
                        value = setFontFamily(value);
                    }
                });
                page.EnumPageTextHtml(changeTextHtml);
                page.SetQuickStyleDef("font", fontName);
            }

            page.Save();
        }

        public void OnInsertTableClick(IRibbonControl control)
        {
            string index = Regex.Match(GetControlId(control), @"\d+$").Groups[0].Value;
            var table = setting.Tables[int.Parse(index)];
            table.Row = table.Row > 0 ? table.Row : 1;
            table.Column = table.Column > 0 ? table.Column : 1;
            table.ForeColor = Regex.Match(table.ForeColor, "[0-9a-fA-F]+").Groups[0].Value;
            table.BackColor = Regex.Match(table.BackColor, "[0-9a-fA-F]+").Groups[0].Value;

            wordHandler.CreateTable(table);
            // 粘贴并去掉末尾的空格
            SendKeys.SendWait("^(v){BACKSPACE}{BACKSPACE}");
        }

        public void OnOpenSettingFileClick(IRibbonControl control)
        {
            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = setting.FilePath;
            process.StartInfo.Arguments = "notepad";
            process.Start();
        }
    }
}
