using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using pp = Microsoft.Office.Interop.PowerPoint;

namespace SlideShowHistory
{
    public class PowerPoint : IDisposable
    {
        private pp.Application powerpointInstance;

        private Timer screenshotTimer;

        private Timer isActiveTimer;

        private Dictionary<int, Image> slideScreenshots;

        private List<Image> screenshotList;

        private List<SlideshowHistoryDialog> historyDialogs;

        private int screenCount;

        private int currentScreenIndex = 1;

        public event EventHandler<PowerPointStatus> StatusChanged;

        public enum PowerPointStatus
        {
            CONNECTED, DISCONNECTED
        }

        public PowerPoint(int screens)
        {
            this.screenCount = screens;
            slideScreenshots = new Dictionary<int, Image>();
            screenshotList = new List<Image>();
            historyDialogs = new List<SlideshowHistoryDialog>();

            // initialize screenshot timer
            screenshotTimer = new Timer();
            screenshotTimer.Interval = 500;
            screenshotTimer.Elapsed += ScreenshotCapture;

            isActiveTimer = new Timer();
            isActiveTimer.Interval = 1000;
            isActiveTimer.Elapsed += IsActiveTimer_Elapsed;
            isActiveTimer.Enabled = true;

            updateHistoryDialogs();
        }

        private void IsActiveTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (powerpointInstance == null)
            {
                InitializePowerpoint();
            }

            try
            {
                var currentApp = powerpointInstance.Active;

            }
            catch (Exception)
            {
                screenshotTimer.Enabled = false;
                powerpointInstance = null;
                OnStatusChanged(PowerPointStatus.DISCONNECTED);
            }
        }

        protected virtual void OnStatusChanged(PowerPointStatus e)
        {
            StatusChanged(this, e);
        }

        private void updateHistoryDialogs()
        {
            // initialized?
            if (historyDialogs.Count == 0)
            {
                for (int i = 0; i < screenCount; i++)
                {
                    // create new screens for history function
                    var dialog = new SlideshowHistoryDialog();
                    dialog.Show();

                    // calculate positions
                    

                    historyDialogs.Add(dialog);
                }
            }

            // update screens
            for (int i = 0; i < screenCount; i++)
            {
                if (screenshotList.Count > i)
                {
                    try
                    {
                        var old = historyDialogs[i].BackgroundImage;
                        historyDialogs[i].BackgroundImage = new Bitmap(screenshotList[i]);

                        if (old != null)
                            old.Dispose();
                    }
                    catch (Exception)
                    {

                    }
                }
            }
        }

        private void ScreenshotCapture(object sender, ElapsedEventArgs e)
        {
            // active connection?
            if (powerpointInstance == null)
            {
                screenshotTimer.Enabled = false;
                return;
            }

            try
            {
                int slideIndex = powerpointInstance.SlideShowWindows[1].View.Slide.SlideIndex;

                if (slideScreenshots.ContainsKey(slideIndex))
                {
                    Image previous = slideScreenshots[slideIndex];
                    slideScreenshots[slideIndex] = Screenshot.CreateScreenshot();

                    previous.Dispose();
                }
                else
                {
                    slideScreenshots[slideIndex] = Screenshot.CreateScreenshot();
                }
            }
            catch (InvalidComObjectException ex)
            {
                screenshotTimer.Enabled = false;
                OnStatusChanged(PowerPointStatus.DISCONNECTED);
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2147467262)
                {
                    screenshotTimer.Enabled = false;
                    OnStatusChanged(PowerPointStatus.DISCONNECTED);
                }
            }
        }

        public bool InitializePowerpoint()
        {
            try
            {
                // remove old instance
                CleanCom();

                // connect to powerpoint
                powerpointInstance = Marshal.GetActiveObject("PowerPoint.Application") as pp.Application;

                if (powerpointInstance != null)
                {
                    powerpointInstance.SlideShowBegin += Powerpoint_SlideShowBegin;
                    powerpointInstance.SlideShowEnd += Powerpoint_SlideShowEnd;
                    powerpointInstance.SlideShowNextSlide += Powerpoint_SlideShowNextSlide;

                    OnStatusChanged(PowerPointStatus.CONNECTED);
                    return true;
                }
            }
            catch (Exception ex)
            {
                // do nothing :(
            }

            OnStatusChanged(PowerPointStatus.DISCONNECTED);
            powerpointInstance = null;

            return false;
        }

        private void Powerpoint_SlideShowEnd(pp.Presentation Pres)
        {
            screenshotTimer.Enabled = false;
        }

        private void Powerpoint_SlideShowBegin(pp.SlideShowWindow Wn)
        {
            slideScreenshots.Clear();
            screenshotTimer.Enabled = true;
            currentScreenIndex = 1;
        }

        private void Powerpoint_SlideShowNextSlide(pp.SlideShowWindow Wn)
        {
            // get current new slide index
            var index = Wn.View.Slide.SlideIndex;
            Image screen;

            var previousIndex = currentScreenIndex > index ? index + 1 : index - 1;
            currentScreenIndex = index;

            if (slideScreenshots.TryGetValue(previousIndex, out screen))
            {
                screenshotList.Insert(0, screen);
                if (screenshotList.Count > screenCount)
                {
                    screenshotList[screenshotList.Count - 1].Dispose();
                    screenshotList.RemoveAt(screenshotList.Count - 1);
                }
            }

            updateHistoryDialogs();
        }

        public void Dispose()
        {
            CleanCom();
        }

        private void CleanCom()
        {
            if (powerpointInstance != null)
            {
                try
                {
                    int refvalue = 0;
                    do
                    {
                        refvalue = Marshal.ReleaseComObject(powerpointInstance);
                    } while (refvalue > 0);
                }
                catch (Exception ex)
                {

                }
            }
        }
    }
}
