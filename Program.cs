using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToolReport
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool ownMutex;
            using (Mutex mutex = new Mutex(true, "toolReport", out ownMutex))
            {
                if (ownMutex)
                {
                    Application.Run(new Form1());
                }
                else
                {
                    MessageBox.Show("Tool is running");
                    Application.Exit();
                }
            }
        }
    }
}
