using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Reflection;

namespace OneNoteAddin.Handler
{
    /// <summary>
    /// visual studio code 的处理类
    /// TODO : 延时在不同机器上会有所不同, 要看一下怎么修改
    /// </summary>
    public class VSCodeHandler
    {
        public WindowHandler windowHandler;

        private string vsCodePath;

        private string filePath;

        public VSCodeHandler(string vsCodePath)
        {
            this.vsCodePath = vsCodePath;
        }

        /// <summary>
        /// 新建一个vscode, 设置格式, 然后隐藏, 每次复制的时候都会
        /// 粘贴到这个里面的编辑器中
        /// </summary>
        /// <param name="codePath">code.exe 的路径</param>
        public void InitialForFormat()
        {
            var pid = Process.GetCurrentProcess().Id;
            StartProcess("OneNoteAddin-CodeStyle-" + pid);

            windowHandler.ShowWindow(false);
        }

        private void StartProcess(string title)
        {
            windowHandler = new WindowHandler();

            filePath = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)),
                title);
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }


            Process vscp = new Process();
            vscp.StartInfo.FileName = vsCodePath;
            vscp.StartInfo.Arguments = "-n \"" + filePath + "\"";
            vscp.Start();

            // 等待visual studio code 启动, 不然标题不会显示文件名
            while (!windowHandler.SetWindow(
                "Chrome_WidgetWin_1",
                title + " - Visual Studio Code"))
            {
                Thread.Sleep(500);
            }
        }

        /// <summary>
        /// 设置编辑器的代码格式
        /// </summary>
        /// <param name="codeStyle">代码格式</param>
        public void ChangeLanguageMode(string codeStyle)
        {
            windowHandler.ActiveWindow();

            // 避免输入法影响
            bool isCapsPressed = WindowHandler.IsCapsLockPressed();
            if (!isCapsPressed)
            {
                WindowHandler.SendCapsLock();
            }
            SendKeys.SendWait("^(k)m");
            Thread.Sleep(300);
            SendKeys.SendWait($"{codeStyle}\n");

            if (!isCapsPressed)
            {
                WindowHandler.SendCapsLock();
            }

            if (!windowHandler.IsWindowVisible())
            {
                windowHandler.DeactiveWindow();
            }
        }

        /// <summary>
        /// 关闭编辑器.
        /// </summary>
        public void Close()
        {
            if (windowHandler != null && !windowHandler.IsWindowClosed())
            {
                windowHandler.ActiveWindow();
                SendKeys.SendWait("^(s)");
                Thread.Sleep(500);
                SendKeys.SendWait("^(w)");
                Thread.Sleep(500);
                windowHandler.CloseWindow();
                windowHandler = null;
                File.Delete(filePath);
            }
        }

        /// <summary>
        /// 从剪切板粘贴, 格式化, 然后复制到剪切板
        /// </summary>
        public void PasteFormatCut()
        {
            windowHandler.ActiveWindow();

            Thread.Sleep(200);
            SendKeys.SendWait("^(v)+%(f)");
            Thread.Sleep(300);
            SendKeys.SendWait("^(ax)");

            windowHandler.DeactiveWindow();
        }
        /// <summary>
        /// 从剪切板粘贴然后保存, 然后编辑代码, 关闭后和保存的文件进行对比,
        /// 如果文件改变了就复制到剪切板中
        /// </summary>
        /// <param name="codePath">code.exe 的路径</param>
        /// <param name="newText">更改后的文本</param>
        /// <returns>文本是否被修改过</returns>
        public bool EditCode(string codeStyle, out string newText)
        {
            // 初始化
            var pid = Process.GetCurrentProcess().Id;
            StartProcess("OneNoteAddinCodeEdit-" + pid);

            // 设置编辑器文件
            windowHandler.ActiveWindow();
            SendKeys.SendWait("^(vs)");
            while (!File.Exists(filePath))
            {
                Thread.Sleep(200);
            }
            string codeText = File.ReadAllText(filePath);
            ChangeLanguageMode(codeStyle);

            // 等待编辑完成(关闭窗口)
            while (!windowHandler.IsWindowClosed())
            {
                Thread.Sleep(500);
            }

            // 查看文件是否变化
            newText = File.ReadAllText(filePath);
            File.Delete(filePath);
            return codeText != newText;
        }
    }
}
