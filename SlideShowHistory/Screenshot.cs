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
            Bitmap printscreen = new Bitmap(Screen.AllScreens[1].Bounds.Width, Screen.AllScreens[1].Bounds.Height);
            using (Graphics graphics = Graphics.FromImage(printscreen as Image))
            {
                graphics.CopyFromScreen(Screen.AllScreens[1].Bounds.X, Screen.AllScreens[1].Bounds.Y, 0, 0, printscreen.Size);
                return printscreen;
            }
        }
    }
}
