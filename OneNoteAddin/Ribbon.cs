using Extensibility;
using Microsoft.Office.Core;
using System;
using System.Runtime.InteropServices;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNoteAddin
{
    /// <summary>
    /// ribbon on menu bar
    /// </summary>
    [Guid("7C5FA097-B7EF-4B5B-93D0-8F5C71070876"), ProgId("OneNoteAddin")]
    public partial class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        private OneNote.Application app;

        private IRibbonUI ribbon;

        #region start

        public void OnConnection(object app, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            this.app = app as OneNote.Application;
            System.Diagnostics.Debugger.Launch();
        }

        public void OnAddInsUpdate(ref Array custom)
        {

        }

        public void Ribbon_Load(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
        }

        public void OnStartupComplete(ref Array custom)
        {
            vsCodeHandler = new Handler.VSCodeHandler(setting.VSCode);
        }

        #endregion

        #region shutdown

        public void OnBeginShutdown(ref Array custom)
        {
            Close();
        }
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Close();
        }

        private void Close()
        {
            if (app != null)
            {
                showForm.Close();
                wordHandler.Close();
                SaveSetting();
                vsCodeHandler.Close();

                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion
    }
}
