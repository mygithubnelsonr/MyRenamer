using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyRenamer
{
    public partial class Splash : Form
    {
        int elapsedtime = 0;
        int timeout=0;

        public Splash()
        {
            InitializeComponent();
        }

        public int Timeout
        {
            set
            {
                timeout = value;
                if (timeout > 0)
                    this.timer1.Enabled = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            elapsedtime++;
            if (elapsedtime > timeout)
                this.Close();
        }

		private void Splash_Load(object sender, EventArgs e)
		{
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.labelVersion.Text = Properties.Settings.Default.Version;
		}

        private void Splash_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Splash_MouseMove(object sender, MouseEventArgs e)
        {
            this.Close();
        }
    }
}
