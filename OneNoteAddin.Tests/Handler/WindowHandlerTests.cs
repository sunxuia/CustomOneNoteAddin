using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteAddin.Handler;
using System;

namespace OneNoteAddin.Tests.Handler
{
    [TestClass]
    public class WindowHandlerTests
    {
        [TestMethod]
        public void TestCloseWindow()
        {
            WindowHandler handler = new WindowHandler();
            handler.Handle = (IntPtr) 0x120c8;
            handler.CloseWindow();
        }
    }
}
