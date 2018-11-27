using log4net;
using log4net.Config;
using log4net.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SlideShowHistory
{
    static class Program
    {
        private static ILog logger = LogManager.GetLogger(typeof(Program));

        public static PowerPoint pp;

        public static NotifyIcon notifyIcon;

        public static ContextMenu contextMenu;

        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.ThreadException += Application_ThreadException;

            XmlConfigurator.ConfigureAndWatch(new System.IO.FileInfo("log4net.xml"));

            // get screens
            int screenCount = Screen.AllScreens.Count();
            if (screenCount < 2)
                return;

            // init screen
            pp = new PowerPoint(screenCount - 2);
            pp.StatusChanged += Pp_StatusChanged;

            // init notify icon
            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = SlideShowHistory.Properties.Resources.pp_history;
            notifyIcon.Visible = true;
            notifyIcon.Text = "PowerPoint History";

            contextMenu = new ContextMenu();
            contextMenu.MenuItems.Add(new MenuItem("Connect to PowerPoint", (o, a) =>
            {
                ConnectToPowerPoint();
            }));
            contextMenu.MenuItems.Add(new MenuItem("&Close", (o, a) =>
            {
                notifyIcon.Visible = false;

                pp.Dispose();
                Application.Exit();
            }));

            notifyIcon.ContextMenu = contextMenu;

            Application.Run();
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            logger.Error("Application ThreadException", e.Exception);
        }

        private static void Pp_StatusChanged(object sender, PowerPoint.PowerPointStatus e)
        {
            if (e == PowerPoint.PowerPointStatus.CONNECTED)
            {
                notifyIcon.Icon = Properties.Resources.pp_history_green;
            }
            else
            {
                notifyIcon.Icon = Properties.Resources.pp_history_red;
            }
        }

        public static void ConnectToPowerPoint()
        {
            pp.InitializePowerpoint();
        }

        public static void ShowBalloon(string title, string text, ToolTipIcon icon)
        {
            notifyIcon.ShowBalloonTip(5000, title, text, icon);
        }
    }
}
