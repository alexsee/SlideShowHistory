using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SlideShowHistory
{
    public partial class SlideshowHistoryDialog : Form
    {
        public SlideshowHistoryDialog()
        {
            InitializeComponent();
        }

        private void SlideshowHistoryDialog_Load(object sender, EventArgs e)
        {

        }

        private void SlideshowHistoryDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.pp.Dispose();
            Application.Exit();
        }
    }
}
