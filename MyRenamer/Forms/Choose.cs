using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyRenamer
{
    public partial class Choose : Form
    {
        private string    _logFile;     
        public int LogID = -1;

        Renamer renamer = new Renamer();

        public Choose()
        {
            InitializeComponent();
        }

        private void Choose_Load(object sender, EventArgs e)
        {
            _logFile = renamer.Logfile;

            StreamReader reader;
            reader = new StreamReader(_logFile);

            string line;
            while (!reader.EndOfStream)
            {
                line = reader.ReadLine();

                if (line.IndexOf("source") > 0)
                {
                    string[] ar = line.Split(']');
                    string id = ar[1].Replace("[","");
                    string date = line.Substring(1, 19);
                    ListViewItem item = new ListViewItem(id);
                    item.SubItems.Add(date);
                    this.listView1.Items.Add(item);
                }
            }
            reader.Close();
            reader.Dispose();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogID = -1;
            this.Hide();
        }

        private void takeOverToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                LogID = Convert.ToInt32(listView1.SelectedItems[0].Text);
            }
            catch
            {
                LogID = -1;
                MessageBox.Show("Rename ID", "Falsches Format im Logfile");
            }
            this.Hide();
        }
    }
}
