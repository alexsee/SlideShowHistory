using log4net;
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
        private static ILog logger = LogManager.GetLogger(typeof(PowerPoint));

        private pp.Application powerpointInstance;

        private Timer screenshotTimer;

        private Timer isActiveTimer;

        private Dictionary<int, Image> slideScreenshots;

        private List<Image> screenshotList;

        private List<SlideshowHistoryDialog> historyDialogs;

        private int screenCount;

        private int currentScreenIndex = 1;

        public event EventHandler<PowerPointStatus> StatusChanged;

        private int slideIndex = 0;

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
            isActiveTimer.Interval = 5000;
            isActiveTimer.Elapsed += IsActiveTimer_Elapsed;
            isActiveTimer.Enabled = true;

            updateHistoryDialogs();
        }

        private void IsActiveTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (powerpointInstance == null)
            {
                if (!InitializePowerpoint())
                {
                    screenshotTimer.Enabled = false;
                    powerpointInstance = null;
                    OnStatusChanged(PowerPointStatus.DISCONNECTED);
                    return;
                }
            }

            try
            {
                var currentApp = powerpointInstance.Active;

            }
            catch (Exception ex)
            {
                // powerpoint is busy, so ignore exception message
                if (ex.HResult == -2147417846)
                    return;

                screenshotTimer.Enabled = false;
                powerpointInstance = null;
                OnStatusChanged(PowerPointStatus.DISCONNECTED);

                logger.Error("Failed to check for PowerPoint active.", ex);
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
                logger.Debug("Create new slide show windows.");

                for (int i = 0; i < screenCount; i++)
                {
                    // create new screens for history function
                    var dialog = new SlideshowHistoryDialog();
                    dialog.Show();

                    // calculate positions
                    System.Windows.Forms.Screen screen = Screenshot.GetMonitor(i + 2);

                    var loc = new Point(screen.Bounds.X, screen.Bounds.Y);
                    dialog.Location = loc;
                    dialog.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    dialog.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    dialog.Bounds = screen.Bounds;

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

                logger.Error("Invalid com object during screenshot.", ex);
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2147467262)
                {
                    screenshotTimer.Enabled = false;
                    OnStatusChanged(PowerPointStatus.DISCONNECTED);
                }

                logger.Error("Unexpected exception during screenshot.", ex);
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

                    if (powerpointInstance.Presentations.Count == 0)
                    {
                        Program.ShowBalloon("No active presentation", "Open a PowerPoint presentation and allow editing permissions.", System.Windows.Forms.ToolTipIcon.Info);
                    }

                    OnStatusChanged(PowerPointStatus.CONNECTED);
                    return true;
                }
            }
            catch (Exception ex)
            {
                // do nothing :(
                logger.Error("Error during PowerPoint initialization.", ex);
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
            logger.Debug("New slide show started.");

            slideScreenshots.Clear();
            screenshotTimer.Enabled = true;
            currentScreenIndex = 1;
        }

        private void Powerpoint_SlideShowNextSlide(pp.SlideShowWindow Wn)
        {
            try
            {
                slideIndex = Wn.View.Slide.SlideIndex;

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
            }
            catch (Exception ex)
            {
                // do nothing
                logger.Error("Exception during next slide.", ex);
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
