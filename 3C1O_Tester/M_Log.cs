using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Configuration;

namespace _3C1O_Tester
{

    public class M_Log
    {
        public enum LogType
        {
            Error,
            Info,
            Log
        }
        string[] log_type = new string[3] { "ERROR", "INFO", "LOG" };
        public string name;
        public string dir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\MuLog\\";
        public bool NoError = true;

        /// <summary>
        /// 
        /// </summary>
        public void Open_Log()
        {
            System.Diagnostics.Process.Start("explorer.exe", dir);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type">记录类型</param>
        /// <param name="str">记录内容</param>
        public void Out_Log(LogType type, string str)
        {
            Out_Log(type, str, DateTime.Now);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type">记录类型</param>
        /// <param name="str">记录内容</param>
        /// <param name="time">记录时间</param>
        public void Out_Log(LogType type, string str, DateTime time)
        {
            Out_Log(log_type[(int)type], str, time);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type">记录类型</param>
        /// <param name="str">记录内容</param>
        public void Out_Log(string type, string str)
        {
            Out_Log(type, str, DateTime.Now);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type">记录类型</param>
        /// <param name="str">记录内容</param>
        /// <param name="time">记录时间</param>
        public void Out_Log(string type, string str, DateTime time)
        {
            if (!NoError) return;
            if (name == null || name == "")
                name = System.Windows.Forms.Application.ExecutablePath.Replace(System.Windows.Forms.Application.StartupPath, "");
            string path = dir + name + "\\" + type + "\\";
            string file = time.Year.ToString() + "_" + time.Month.ToString() + "_" + time.Day.ToString() + ".txt";
            if (!File.Exists(path + file))
            {
                Directory.CreateDirectory(path);
                File.Create(path + file).Close();
            }
            StreamWriter sw = File.AppendText(path + file);
            sw.WriteLine(time.ToString() + ":" + time.Millisecond.ToString("000") + " " + str);
            sw.Flush();
            sw.Close();
        }

    }
}
