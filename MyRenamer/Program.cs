using System;
using System.Windows.Forms;
using System.Threading;

namespace MyRenamer
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            const string appName = "MyRenamer";
            bool createNew;

            mutex = new Mutex(true, appName, out createNew);

            if (!createNew)
            {
                //app is already running! Exiting the application
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new Renamer());
        }
    }
}
