using Extensibility;
using Microsoft.Office.Core;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNoteAddin
{
    [Guid("7C5FA097-B7EF-4B5B-93D0-8F5C71070876"), ProgId("OneNoteAddin")]
    public partial class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        private OneNote.Application app;
        
        public void OnConnection(object app, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            this.app = app as OneNote.Application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Close();
        }

        private void Close()
        {
            if (app != null)
            {
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {

        }

        public void OnStartupComplete(ref Array custom)
        {

        }

        public void OnBeginShutdown(ref Array custom)
        {
            Close();
        }
        
        public string GetCustomUI(string ribbonID)
        {
            return Properties.Resources.Ribbon;
        }

        public IStream GetImage(string imageName)
        {
            MemoryStream imageStream = new MemoryStream();
            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            var prop = typeof(Properties.Resources).GetProperty(imageName ?? "", flags);
            if (prop != null)
            {
                (prop.GetValue(null, null) as Image).Save(imageStream, ImageFormat.Png);
            }
            else
            {
                Properties.Resources.DefaultImage.Save(imageStream, ImageFormat.Png);
            }
            return new CCOMStreamWrapper(imageStream);
        }
    }
}
