using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SlideShowHistory
{
    public class Screenshot
    {
        public static Image CreateScreenshot()
        {
            Screen screen = GetMonitor(1);

            Bitmap printscreen = new Bitmap(screen.Bounds.Width, screen.Bounds.Height);
            using (Graphics graphics = Graphics.FromImage(printscreen as Image))
            {
                graphics.CopyFromScreen(screen.Bounds.X, screen.Bounds.Y, 0, 0, printscreen.Size);
                return printscreen;
            }
        }

        public static Screen GetMonitor(int monitorNumber)
        {
            int X = 0;
            Screen currentScreen = null;

            for (int i = 0; i <= monitorNumber; i++)
            {
                currentScreen = FindMonitor(X);
                X += currentScreen.WorkingArea.Width;
            }

            return currentScreen;
        }

        public static Screen FindMonitor(int X)
        {
            foreach(Screen screen in Screen.AllScreens)
            {
                if (screen.WorkingArea.X == X)
                    return screen;
            }

            return null;
        }
    }
}
