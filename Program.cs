using System;
using System.Windows.Forms;

namespace Hyper_V_Manager
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
#if DEBUG
            while (!System.Diagnostics.Debugger.IsAttached)
                System.Threading.Thread.Sleep(100);
#endif
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
