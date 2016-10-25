using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace _3C1O_Tester
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool createdNew;
            //系统能够识别有名称的互斥，因此可以使用它禁止应用程序启动两次
            //第二个参数可以设置为产品的名称:Application.ProductName
            //每次启动应用程序，都会验证名称为SingletonWinAppMutex的互斥是否存在
            Mutex mutex = new Mutex(false, Application.ProductName, out createdNew);

            //如果已运行，则在前端显示
            //createdNew == false，说明程序已运行
            if (!createdNew)
            {
                Process instance = GetExistProcess();
                if (instance != null)
                {
                    SetForegroud(instance);
                    Application.Exit();
                    return;
                }
            }


            //if (!File.Exists("CSkin.dll"))
            //{
            //    byte[] res;//创建byte数组，装资源 
            //    res = new byte[Properties.Resources.CSkin_dll.Length];//一开始我没有定义大小，就老报错。所以要确定数组大小。 
            //    Properties.Resources.CSkin_dll.CopyTo(res, 0);//将apk资源导入byte数组中 
            //                                                  //下面就是将byte写入文件了，可以在具体，比如覆盖还是删文件改名什么的 
            //    FileStream fs = new FileStream("CSkin.dll", FileMode.Create, FileAccess.Write);
            //    fs.Write(res, 0, res.Length);
            //    fs.Close();
            //}
            //string path = @"CSkin.dll";
            //File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.System);
            //File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Hello a = new Hello();
            DialogResult b = a.ShowDialog();
            if (b == DialogResult.Yes || b == DialogResult.OK)
            {
                Application.Run(new Form1(b));
            }
        }
        /// <summary>
        /// 查看程序是否已经运行
        /// </summary>
        /// <returns></returns>
        private static Process GetExistProcess()
        {
            Process currentProcess = Process.GetCurrentProcess();
            foreach (Process process in Process.GetProcessesByName(currentProcess.ProcessName))
            {
                if ((process.Id != currentProcess.Id) &&
                    (Assembly.GetExecutingAssembly().Location == currentProcess.MainModule.FileName))
                {
                    return process;
                }
            }
            return null;
        }
        /// <summary>
        /// 使程序前端显示
        /// </summary>
        /// <param name="instance"></param>
        private static void SetForegroud(Process instance)
        {
            IntPtr mainFormHandle = instance.MainWindowHandle;
            if (mainFormHandle != IntPtr.Zero)
            {
                ShowWindowAsync(mainFormHandle, 1);
                SetForegroundWindow(mainFormHandle);
            }
        }
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);
    }
}
