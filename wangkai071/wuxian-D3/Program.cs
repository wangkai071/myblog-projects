using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wuxian
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            Process[] tProcess = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (tProcess.Length > 1)
            {
                Application.Exit();
            }
            else
            {
                Application.Run(new Form1());
            }

        }
    }
}
