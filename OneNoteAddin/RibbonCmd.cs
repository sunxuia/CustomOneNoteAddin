using Extensibility;
using Microsoft.Office.Core;
using OneNoteAddin.Setting;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNoteAddin
{
    /// <summary>
    /// ribbon events
    /// </summary>
    public partial class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        public void onTextChange(IRibbonControl control, string newText)
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

        public void onButtonClick(IRibbonControl control)
        {
            string id = GetControlId(control);
            MessageBox.Show($"{id} clicked");
        }

        public void OpenXmlFile(IRibbonControl control)
        {
            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = setting.FilePath;
            process.StartInfo.Arguments = "notepad";
            process.Start();
        }
    }
}
