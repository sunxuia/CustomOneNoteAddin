using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteAddin.Handler;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteAddin.Tests.Handler
{
    [TestClass]
    public class VSCodeHandlerTests
    {
        [TestMethod]
        public void TestPastFormatCut()
        {
            VSCodeHandler vsc = new VSCodeHandler("code");
            vsc.InitialForFormat();
            vsc.ChangeLanguageMode("C#");
            vsc.PasteFormatCut();

            vsc.Close();
        }

        [TestMethod]
        public void TestEditCode()
        {
            VSCodeHandler vsc = new VSCodeHandler("c#");
            string newText;
            bool res = vsc.EditCode("code", out newText);
            Debugger.Log(0, "", $"{res}");
        }
    }
}
