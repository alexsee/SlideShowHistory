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
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            using (Graphics graphics = Graphics.FromImage(printscreen as Image))
            {
                graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
                return printscreen;
            }
        }
    }
}
