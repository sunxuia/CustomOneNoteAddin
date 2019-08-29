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
        private WindowHandler windowHandler;

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
            try
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
                do
                {
                    Thread.Sleep(500);
                } while (!windowHandler.SetWindow(
                    "Chrome_WidgetWin_1",
                    title + " - Visual Studio Code"));
                IsInitialed = true;
            }
            catch (Exception err)
            {
                MessageBox.Show("Error while start visual studio code : \n" + err.ToString());
            }
        }

        public bool IsInitialed
        {
            get; private set;
        }

        /// <summary>
        /// 设置编辑器的代码格式
        /// </summary>
        /// <param name="codeStyle">代码格式</param>
        public void ChangeLanguageMode(string codeStyle)
        {
            bool isVisible = windowHandler.IsWindowVisible();
            if (!isVisible)
            {
                windowHandler.ActiveWindow();
            }

            // 避免输入法影响
            bool isCapsPressed = WindowHandler.IsCapsLockPressed();
            if (!isCapsPressed)
            {
                WindowHandler.SendCapsLock();
            }
            SendKeys.SendWait("^(k)m");
            Thread.Sleep(350);
            SendKeys.SendWait($"{codeStyle.ToUpper()}");
            Thread.Sleep(150);
            SendKeys.SendWait("\n");

            Thread.Sleep(150);
            string nowTitle = windowHandler.GetWindowTitle();
            if (nowTitle.Contains("settings.json"))
            {
                SendKeys.SendWait("^(w)");
                MessageBox.Show(codeStyle + " configuration not exist!");
            }

            if (!isCapsPressed)
            {
                WindowHandler.SendCapsLock();
            }

            if (!isVisible)
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

            // 粘贴
            SendKeys.SendWait("^(v)");
            // 格式化
            SendKeys.SendWait("+%(f)");
            // 等待格式化
            Thread.Sleep(150);
            // 发送alt, 避免调出菜单的错误
            SendKeys.SendWait("%");
            Thread.Sleep(150);
            // 全选并剪切
            SendKeys.SendWait("^(ax)");

            windowHandler.DeactiveWindow();
        }

        /// <summary>
        /// 从剪切板粘贴然后复制
        /// </summary>
        public void PasteCut()
        {
            windowHandler.ActiveWindow();

            // 粘贴
            SendKeys.SendWait("^(v)");
            // wait for rendering
            Thread.Sleep(150);
            // 全选并剪切
            SendKeys.SendWait("^(ax)");

            windowHandler.DeactiveWindow();
        }

        /// <summary>
        /// 从剪切板粘贴然后保存, 然后编辑代码, 关闭后和保存的文件进行对比,
        /// 如果文件改变了就复制到剪切板中
        /// </summary>
        /// <param name="codePath">code.exe 的路径</param>
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
            string codeText;
            do
            {
                Thread.Sleep(100);
            } while (!TryReadFile(filePath, out codeText));

            Thread.Sleep(500);
            ChangeLanguageMode(codeStyle);

            // 等待编辑完成(关闭窗口)
            while (!windowHandler.IsWindowClosed())
            {
                Thread.Sleep(500);
            }

            // 读取修改后的文件
            while (!TryReadFile(filePath, out newText))
            {
                Thread.Sleep(100);
            }
            File.Delete(filePath);
            return codeText != newText;
        }

        private bool TryReadFile(string filePath, out string text)
        {
            try
            {
                text = File.ReadAllText(filePath);
                return true;
            }
            catch (Exception err)
            {
                text = err.ToString();
                return false;
            }
        }

        /// <summary>
        /// 隐藏/ 显示vsc
        /// </summary>
        /// <param name="show">是否显示</param>
        public void ShowWindow(bool show)
        {
            windowHandler.ShowWindow(show);
            if (show)
            {
                windowHandler.ActiveWindow();
            }
        }

        /// <summary>
        /// 窗口是否是可见的
        /// </summary>
        public bool IsWindowVisible
        {
            get
            {
                return windowHandler.IsWindowVisible();
            }
        }
    }
}
