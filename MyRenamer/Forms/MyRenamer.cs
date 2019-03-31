using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using WMPLib;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using NRSoft.FunctionPool;

/// <summary>
/// History:
/// 27.01.2019  2.1.4   Bugfix, add error handling in Action_File_FindItem, return int -1 if file not found
/// </summary>

// ToDo weiter geht's mit 
// 

namespace MyRenamer
{
    #region class Renamer
    public partial class Renamer : Form
	{
        #region private fields
        private const string _company = "NrSoft";
		private const string _productKey = "MyRenamer.net";
        private string _version = Properties.Settings.Default.Version;
        private const string _logFile = "MyRenamer.log";
        private bool showSplash = false;
		private ListViewItemComparer listViewItemComparer = new ListViewItemComparer();
        private WindowsMediaPlayer player = new WindowsMediaPlayer();
        private RegistryH rh = new RegistryH();
        private GeneralH gh = new GeneralH();
        private FileSystemUtils fh = new FileSystemUtils();
        private string _currentPlaying = "";
        private string _sourcePath = "";
        private string _keyRoot = "Software\\" + _company + "\\" + _productKey;
        private int _logID = -1;
        private Int64 newID;
        #endregion

        #region properties
        public string Company
        {
            get { return _company; }
        }

        public string ProductKey
        {
            get { return _productKey; }
        }

        public string Version
        {
            get { return _version; }
        }

        public string SourcePath
        {
            set { _sourcePath = value; }
        }

        public string Logfile
        {
            get { return _logFile; }
        }

        public int LogID
        {
            get { return _logID; }
            set { _logID = value; }
        }


        #endregion

        #region CTOR
        public Renamer()
        {
            InitializeComponent();
        }
        #endregion

        #region timer events
        private void timerUnselect_Tick(object sender, EventArgs e)
        {
            if( _currentPlaying == "")
            {
                _currentPlaying = player.currentMedia.sourceURL;
            }

            if (_currentPlaying != player.currentMedia.sourceURL)
            {
                _currentPlaying = player.currentMedia.sourceURL;
            }

            try
            { 
                int itm = Action_File_FindItem(_currentPlaying);

                listViewFiles.Items[itm].Selected = false;
                listViewFiles.Items[itm].BackColor = Color.GreenYellow;

                decimal m = (decimal)trackBarPosition.Maximum / 100;
                decimal v = (decimal)trackBarPosition.Value;
                int percent = (int)(v / m);

                string s = player.controls.currentPositionString;
                string text = listViewFiles.Items[itm].SubItems[1].Text;

                toolTipPosition.SetToolTip(trackBarPosition, "playing \n" + text + "\n" + s + "\t(" + percent + "%)\nVolume=" + v.ToString());
            }
            catch (Exception ex)
            {
                Debug.Assert(true, ex.Message);
            }
        }

        private void timerMusic_Tick(object sender, EventArgs e)
        {
            if (player.playState == WMPPlayState.wmppsPlaying)
            {
                double dblPosition = player.controls.currentPosition;
                double dblDuration = player.currentMedia.duration;

                trackBarPosition.Maximum = (int)dblDuration;
                toolStripProgressBar.Maximum = (int)dblDuration;

                if (trackBarPosition.Maximum > dblPosition)
                {
                    trackBarPosition.Value = (int)dblPosition;
                }
            }
            else if(player.playState == WMPPlayState.wmppsStopped)
            {
                timerMusic.Enabled = false;
                player.controls.stop();
                timerMusic.Stop();
                trackBarPosition.Value = 0;
            }
            else if(player.playState == WMPPlayState.wmppsPaused)
            {
                player.controls.play();
            }
        }
        #endregion

        #region Form Events
        private void Renamer_Load(object sender, EventArgs e)
        {
            // add eventhandler
            player.PlayStateChange += new WMPLib._WMPOCXEvents_PlayStateChangeEventHandler(Player_PlayStateChange);
            player.MediaError += new WMPLib._WMPOCXEvents_MediaErrorEventHandler(Player_MediaError);

            if (showSplash == true)
            {
                Splash f = new Splash();
                Application.DoEvents();
                f.Timeout = 5;
                f.Show(this);
            }

            RestoreSettings();

            InitListViews();

            player.settings.volume = trackBarVolume.Value;
            FillCombo(this.ComboBoxFiles, "Mp3DirList");
            FillCombo(this.ComboBoxTitles, "TextFileList");
        }

        private void Renamer_FormClosing(object sender, FormClosingEventArgs e)
        {
            string strTag = menuMainFileExit.Tag.ToString();

            WMPPlayState playerState = player.playState;

            timerMusic.Stop();
            player.controls.pause();

            if (e.Cancel == false)
            {
                player.close();
                player = null;

                SaveSettings();

            }
        }

        private void SplitContainer1_DoubleClick(object sender, EventArgs e)
        {
            SplitContainer1.SplitterDistance = this.Width / 2 - SplitContainer1.SplitterWidth;
        }

        private void listViewFiles_ItemActivate(object sender, EventArgs e)
        {
            foreach (ListViewItem itmX in listViewNames.SelectedItems)
            {
                itmX.BackColor = Color.FromArgb(170, 255, 170);
            }
        }

        private void listViewFiles_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            foreach (ListViewItem itmX in listViewNames.SelectedItems)
            {
                itmX.BackColor = Color.FromArgb(170, 255, 170);
            }
        }

        private void listViewFiles_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                listViewFiles.Items.Clear();

                string[] filenames = (string[])e.Data.GetData(DataFormats.FileDrop, false);

                DirectoryInfo di;
                di = new DirectoryInfo(filenames[0]);
                string strFolder;

                if ((di.Attributes & FileAttributes.Directory) > 0)
                    strFolder = di.FullName;
                else
                    strFolder = di.Parent.FullName;

                Add2combo(ComboBoxFiles, strFolder);
                ComboBoxFiles.Text = strFolder;
            }
        }

        private void listViewFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void ListViewNames_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop, false);
                string fileName = fileNames[0];

                listViewNames.Items.Clear();
                Add2combo(ComboBoxTitles, fileName);
                ComboBoxTitles.Text = fileName;
            }
        }

        private void ListViewNames_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void ContextMenuFilesClear_Click(object sender, EventArgs e)
        {
            listViewFiles.Items.Clear();
            this.ComboBoxFiles.Text = "";
        }

        private void ComboBoxFiles_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuComboFiles.Show((ComboBox)sender, e.Location);
            }
        }

        private void removeItemToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void listViewFiles_DoubleClick(object sender, EventArgs e)
        {
            if (listViewFiles.SelectedItems.Count > 0)
            {
                buttonPlay.PerformClick();
            }
        }

        private void trackBarPosition_Scroll(object sender, EventArgs e)
        {
            player.controls.currentPosition = trackBarPosition.Value;
        }

        private void trackBarVolume_Scroll(object sender, EventArgs e)
        {
            player.settings.volume = trackBarVolume.Value;
        }

        private void trackBarVolume_MouseUp(object sender, MouseEventArgs e)
        {
            listViewFiles.Focus();
        }

        private void trackBarPosition_MouseUp(object sender, MouseEventArgs e)
        {
            listViewFiles.Focus();
        }

        private void trackBarPosition_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                toolStripProgressBar.Value = trackBarPosition.Value;
            }
            catch (Exception ex)
            {
                Debug.Assert(true, ex.Message);
            }
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Action_File_SelectAll();
        }

        private void playToolStripMenuItem_Click(object sender, EventArgs e)
        {
            buttonPlay.PerformClick();
        }

        #endregion Form Events

        #region Listview Events
        private void ListViewNames_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete)
			    Action_Cut(listViewNames);

            if (e.KeyCode == Keys.F2)
                Action_Text_Edit();
            
            if (e.KeyValue == 107)
                Action_MoveUp(listViewNames);

            if (e.KeyValue == 109)
                Action_MoveDown(listViewNames);
		}

        private void listViewNames_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuNames.Show((ListView)sender, e.Location);
            }
        }

        private void listViewNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            statusStripNameselected.Text = listViewNames.SelectedItems.Count.ToString();
        }

		private void listViewFiles_SelectedIndexChanged(object sender, EventArgs e)
		{
            statusStripFileselected.Text = listViewFiles.SelectedItems.Count.ToString();
		}

		/// <summary>
		/// Handles the ColumnClick event of the listViewFiles control.
		/// </summary>
		/// <param name="sender">The source of the event.</param>
		/// <param name="e">The <see cref="System.Windows.Forms.ColumnClickEventArgs"/> instance containing the event data.</param>
		private void listViewFiles_ColumnClick(object sender, ColumnClickEventArgs e)
		{
			SortOrder sort_order;

			// aktuelle Sortierung
			// Umsortieren der Liste beim Klick auf eine Spalte

			int curColumn = e.Column;

			if (curColumn == this.listViewItemComparer.CurrentColumn)
			{
				sort_order = listViewFiles.Sorting;
				SortOrder old_order = listViewItemComparer.SortOrder;

				if (sort_order == SortOrder.Ascending)
					sort_order = SortOrder.Descending;
				else
					sort_order = SortOrder.Ascending;

				listViewFiles.Sorting = sort_order;
				listViewFiles.Sort();

				this.listViewFiles.ListViewItemSorter = new ListViewItemComparer(e.Column);
				
				
			}
			else
			{
				sort_order = SortOrder.Ascending;
				
				listViewFiles.Sorting = sort_order;

				this.listViewFiles.ListViewItemSorter = new ListViewItemComparer(e.Column);

				listViewFiles.Sort();
			}

			this.listViewItemComparer.CurrentColumn = curColumn;
			this.listViewItemComparer.SortOrder = sort_order;
		}

        private void listViewFiles_KeyDown(object sender, KeyEventArgs e)
        {
            Console.WriteLine(e.KeyCode);

			if (e.KeyCode == Keys.Delete)
			{
                this.Action_Cut(this.listViewFiles);
			}
            if (e.KeyValue == 107)
            {
                this.Action_MoveUp(this.listViewFiles);
            }
            if (e.KeyValue == 109)
            {
                this.Action_MoveDown(this.listViewFiles);
            }
        }

        #endregion Listview Events

        #region ComboBox Events
        private void ComboBoxNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxTitles.Text.Trim().Length > 0)
            {
                if (!FormsUtilities.IsMenuChecked(menuMainTextformat))
                {
                    MessageBox.Show(new Form { TopMost = true }, "Es muss ein Textformat ausgewählt werden!", "Renamer - OpenText",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ComboBoxTitles.Text = string.Empty;
                    return;
                }
                PrepareNameList(ComboBoxTitles.Text);
            }
            else
            {
                listViewNames.Items.Clear();
            }
        }

        private void ComboBoxNames_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void ComboBoxNames_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                listViewNames.Items.Clear();

                string[] TitlesNames = (string[])e.Data.GetData(DataFormats.FileDrop, false);

                DirectoryInfo di;
                di = new DirectoryInfo(TitlesNames[0]);

                string strFile;

                if ((di.Attributes & FileAttributes.Directory) > 0)
                    strFile = di.FullName;
                else
                    strFile = di.FullName;

                Add2combo(ComboBoxTitles, strFile);
                ComboBoxTitles.Text = strFile;
            }
        }

        private void ComboBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxFiles.Text.Trim().Length > 0)
            {
                DirectoryInfo di;
                di = new DirectoryInfo(ComboBoxFiles.Text);

                if (!di.Exists)
                {
                    DialogResult result;
                    result = MessageBox.Show("Folder not found! Delete entry?", "Renamer - Files", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if (result == DialogResult.Yes)
                    {
                        ComboBoxFiles.Items.Remove(ComboBoxFiles.SelectedItem);
                    }
                }
            }
            else
            {
                listViewFiles.Items.Clear();
                statusStripFilecount.Text = " 0";
                statusStripFileselected.Text = " 0";
            }
            PrepareFileList(ComboBoxFiles.Text);
            listViewFiles.Focus();
        }

        private void ComboBoxFiles_DragEnter(object sender, DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effect = DragDropEffects.Copy;
			}
			else
			{
				e.Effect = DragDropEffects.None;
			}
		}

		private void ComboBoxFiles_DragDrop(object sender, DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				listViewFiles.Items.Clear();

				string[] filenames = (string[]) e.Data.GetData(DataFormats.FileDrop, false);

				DirectoryInfo di;
				di = new DirectoryInfo(filenames[0]);
				string strFolder;

				if ((di.Attributes & FileAttributes.Directory) > 0)
					strFolder = di.FullName;
				else
					strFolder = di.Parent.FullName;

                Add2combo(ComboBoxFiles, strFolder);
				ComboBoxFiles.Text = strFolder;
			}
		}

		private void ComboBoxFiles_KeyPress(object sender, KeyPressEventArgs e)
		{
			PrepareFileList(ComboBoxFiles.Text);
		}

		private void ComboBoxFiles_TextChanged(object sender, EventArgs e)
		{
            player.controls.stop();
            timerMusic.Stop();
            trackBarPosition.Value = 0;

            string path = ComboBoxFiles.Text;
            ComboBoxFiles.BackColor = Color.PaleTurquoise;

            if (String.IsNullOrEmpty(path))    // && Directory.Exists(path) == false)
            {
                listViewFiles.Items.Clear();
            }
            else
            {
                if (Directory.Exists(path) == true)
                {
                    ComboBoxFiles.BackColor = Color.PaleTurquoise;
                    //PrepareFileList(ComboBoxFiles.Text);
                }
                else
                {
                    ComboBoxFiles.BackColor = Color.LightPink;
                }
            }
        }
        
        #endregion ComboBox Events

        #region Menu Events
        private void menuMainFileTouchfiles_Click(object sender, EventArgs e)
        {
            this.Action_File_TouchFiles();
        }

        private void menuMainFileExit_Click(object sender, EventArgs e)
        {
            menuMainFileExit.Tag = "1";
            this.Close();
        }

        private void menuMainFileRename_Click(object sender, EventArgs e)
        {
            Action_File_RenameDo();
        }

        private void menuMainExtrasDirectMode_Click(object sender, EventArgs e)
        {
            menuMainExtrasDirectMode.Checked = !menuMainExtrasDirectMode.Checked;
        }

        private void menuMainExtrasPlayBackStop_Click(object sender, EventArgs e)
        {
            buttonStop.PerformClick();
        }

        private void menuMainFileOpenText_Click(object sender, EventArgs e)
        {
            if (!FormsUtilities.IsMenuChecked(menuMainTextformat))
            {
                MessageBox.Show("Es muss ein Textformat ausgewählt werden!", "Renamer - OpenText");
                return;
            }

            string strFileName = "";

            OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|XML files (*.xml)|*.xml|All files (*.*)|*.*";
            OpenFileDialog1.FilterIndex = 2;
            OpenFileDialog1.RestoreDirectory = true;

            if (OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                strFileName = OpenFileDialog1.FileName;
            }

            Add2combo(ComboBoxTitles, strFileName);
            ComboBoxTitles.Text = strFileName.ToLower();
        }

        private void menuMainFileOpenMp3_Click(object sender, EventArgs e)
        {
            string strFolderName = "";

            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                strFolderName = folderBrowserDialog1.SelectedPath;
            }

            Add2combo(ComboBoxFiles, strFolderName);
            ComboBoxFiles.Text = strFolderName.ToLower();
        }

        private void menuMainFileUndo_Click(object sender, EventArgs e)
        {
            Action_File_RenamePrepUndo();
        }

        private void menuMainFleSaveasCD_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewNames, 0);
        }

        private void menuMainFleSaveasDb_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewNames, 1);
        }

        private void menuMainFileSaveasAsIs_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewNames, 2);
        }

        private void menuMainTextformatAsIs_Click(object sender, EventArgs e)
        {
            FormsUtilities.UncheckMenu(menuMainTextformat);
            menuMainTextformatAsIsToolStrip.Checked = true;
        }

        private void menuMainTextformatCD_Click(object sender, EventArgs e)
        {
            FormsUtilities.UncheckMenu(menuMainTextformat);
            menuMainTextformatCDToolStrip.Checked = true;
        }

        private void menuMainTextformatCDAid_Click(object sender, EventArgs e)
        {
            FormsUtilities.UncheckMenu(menuMainTextformat);
            menuMainTextformatCDAidToolStrip.Checked = true;
        }

        private void menuMainTextformatDb_Click(object sender, EventArgs e)
        {
            FormsUtilities.UncheckMenu(menuMainTextformat);
            menuMainTextformatDbToolStrip.Checked = true;
        }

        private void menuMainExtrasTestOnly_Click(object sender, EventArgs e)
        {
            // int itemNr = Action_File_FindItem("Kingston Trio - Tom Dooley");
            // listViewFiles.Items[itemNr].Selected = false;
            
            int renameID = (int)Registry.GetValue(Registry.CurrentUser.Name + "\\" + _keyRoot + "\\Settings", "renameID", - 1);
            var val = rh.GetSetting("Settings", "renameID", "-1");
            renameID = Convert.ToInt16(val);
            Debug.Print(val.ToString());
        }

        private void menuMainExtrasPlayBackPlay_Click(object sender, EventArgs e)
        {
            buttonPlay.PerformClick();
        }

        private void menuMainExtrasPlayBackPause_Click(object sender, EventArgs e)
        {
            buttonPause.PerformClick();
        }

        private void menuMainHelpAbout_Click(object sender, EventArgs e)
        {
            AboutBox f = new AboutBox
            {
                TimeOut = 0 // timeout in sekunden, 0 = kein timeout
            };
            f.Show();
        }

        #endregion Menu Events

        #region ContextMenuNames Events

        private void contextMenuNamesFlip_Click(object sender, EventArgs e)
        {
            Action_Text_Wechseln();
        }

        private void contextMenuNamesEdit_Click(object sender, EventArgs e)
        {

        }

        private void contextMenuNamesCut_Click(object sender, EventArgs e)
        {
            Action_Cut(this.listViewNames);
        }

        private void contextMenuNamesSelectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvItm in this.listViewNames.Items)
            {
                lvItm.Selected = true;
            }
        }

        private void contextMenuNamesClear_Click(object sender, EventArgs e)
        {
            this.listViewNames.Items.Clear();
        }

        private void contextMenuNamesCopyToClipboardAsCD_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewNames, 0);
        }

        private void contextMenuNamesCopyToClipboardAsDb_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewNames, 1);
        }

        private void contextMenuNamesCopyToClipboardAsIs_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewNames, 2);
        }

        private void contextMenuNamesPaste_Click(object sender, EventArgs e)
        {
            Action_Text_Paste();
        }

        private void contextMenuNamesToProperCase_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem li in listViewNames.SelectedItems)
            {
                string strText = FormsUtilities.ToProperCase(li.SubItems[1].Text);
                li.SubItems[1].Text = strText;
            }
        }

        private void contextMenuFilesCut_Click(object sender, EventArgs e)
        {
            this.Action_Cut(this.listViewFiles);
        }
        
        private void contextMenuFilesPaste_Click(object sender, EventArgs e)
        {
            Action_File_Paste();
        }

        private void contextMenuFilesCopyToClipAsIs_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewFiles, 2);
        }

        private void contextMenuFilesCopyToClipCD_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewFiles, 0);
        }

        private void contextMenuFilesCopyToClipDB_Click(object sender, EventArgs e)
        {
            Action_CopyToClip(this.listViewFiles, 1);
        }
        
        #endregion contextMenu Events

        #region player buttons events
        private void buttonPlay_Click(object sender, EventArgs e)
        {
            timerMusic.Stop();
            player.controls.stop();

            if (this.listViewFiles.SelectedItems.Count == 0)
            {
                return;
            }

            IWMPPlaylist playlist = player.playlistCollection.newPlaylist("MyPlaylist");
            IWMPMedia media;

            foreach (ListViewItem itmX in listViewFiles.SelectedItems)
            {
                string song = Path.Combine(ComboBoxFiles.Text, itmX.SubItems[1].Text + itmX.SubItems[2].Text);
                media = player.newMedia(song);
                playlist.appendItem(media);

            }
            player.settings.autoStart = true;
            player.currentPlaylist = playlist;
            timerMusic.Start();
            timerUnselect.Start();
            listViewFiles.Focus();
        }

        private void buttonPlay_MouseEnter(object sender, EventArgs e)
        {
            buttonPlay.BackColor = Color.Gray;
        }

        private void buttonPlay_MouseLeave(object sender, EventArgs e)
        {
            buttonPlay.BackColor = Color.WhiteSmoke;
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            player.controls.stop();
			player.URL = "";
            timerMusic.Stop();
            timerUnselect.Stop();
            trackBarPosition.Value = 0;
            toolStripProgressBar.Value = 0;
        }

        private void buttonStop_MouseEnter(object sender, EventArgs e)
        {
            buttonStop.BackColor = Color.Gray;
        }

        private void buttonStop_MouseLeave(object sender, EventArgs e)
        {
            buttonStop.BackColor = Color.WhiteSmoke;
        }

        private void buttonPause_Click(object sender, EventArgs e)
        {
            if(player.playState == WMPPlayState.wmppsPaused)
            {
                player.controls.play();
                timerMusic.Start();
                return;
            }

            if (player.playState == WMPPlayState.wmppsPlaying)
            {
                player.controls.pause();
                timerMusic.Stop();
                return;
            }
        }

        private void buttonPause_MouseEnter(object sender, EventArgs e)
        {
            buttonPause.BackColor = Color.Gray;
        }

        private void buttonPause_MouseLeave(object sender, EventArgs e)
        {
            buttonPause.BackColor = Color.WhiteSmoke;
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            player.controls.next();
        }

        private void buttonNext_MouseEnter(object sender, EventArgs e)
        {
            buttonNext.BackColor = Color.Gray;
        }

        private void buttonNext_MouseLeave(object sender, EventArgs e)
        {
            buttonNext.BackColor = Color.WhiteSmoke;
        }

        #endregion Player Buttons

        #region public Methodes

        private void InitSettings()
        {
            // root key = hkcu\software

            DateTime dT2 = DateTime.Now;

            rh.SaveSetting("Settings", "Version", _version);
            rh.SaveSetting("Settings", "LastRun", dT2.ToString());
            rh.SaveSetting("Settings", "ForbChars", "\\/:?");

        }

        /// <summary>
        /// initialize from and settings on start
        /// </summary>
        private void RestoreSettings()
        {
            int retVal;

            rh.CompanyName = Company;
            rh.ProductName = ProductName;

            if (rh.GetSetting("Settings", "Version", "") != null)
            {
                InitSettings();
            }

            rh.SaveSetting("Settings", "Version", Version);
            rh.SaveSetting("Settings", "LastRun", DateTime.Now.ToString());

            // restore last settings
            retVal = Convert.ToInt16(rh.GetSetting(@"Settings\Form", "Top", "100"));
            this.Top = retVal <= 0 ? 100 : retVal;
            retVal = Convert.ToInt16(rh.GetSetting(@"Settings\Form", "Left", "100"));
            this.Left = retVal <=0 ? 100 : retVal;
            this.Width = Convert.ToInt16(rh.GetSetting(@"Settings\Form", "Width", "100"));
            this.Height = Convert.ToInt16(rh.GetSetting(@"Settings\Form", "Height", "100"));
            string defaultvalue = (this.Width / 2).ToString();
            retVal = Convert.ToInt16(rh.GetSetting("Settings\\Form", "Splitter", defaultvalue));
            SplitContainer1.SplitterDistance = retVal;
            trackBarVolume.Value = Convert.ToInt16(rh.GetSetting("Settings", "Volume","0"));
        }

        private void SaveSettings()
        {
            string sReg;

            // save form position and size
            int w = this.Size.Width;
            int h = this.Size.Height;
            int t = this.Top;
            int l = this.Left;
            int s = SplitContainer1.SplitterDistance;

            rh.SaveSetting(@"Settings\Form", "Width", w.ToString());
            rh.SaveSetting(@"Settings\Form", "Height", h.ToString());
            rh.SaveSetting(@"Settings\Form", "Top", t.ToString());
            rh.SaveSetting(@"Settings\Form", "Left", l.ToString());
            rh.SaveSetting(@"Settings\Form", "Splitter", s.ToString());
            rh.SaveSetting(@"Settings", "Volume", trackBarVolume.Value.ToString());

            // save file combobox items
            sReg = "";
            List<string> Folders = new List<string>();
            foreach (string element in this.ComboBoxFiles.Items)
            {
                if (element.Trim() != "")
                {
                    Folders.Add(element.ToString());
                }
            }

            if (Folders.Count > 0)
            {
                foreach (string element in Folders)
                {
                    sReg += element + ";";
                }
                rh.SaveSetting(@"Settings\Pfade", "Mp3DirList", sReg);
            }

            // save name combobox items
            sReg = "";
            List<string> Names = new List<string>();
            foreach (string element in this.ComboBoxTitles.Items)
            {
                if (element.Trim() != "")
                {
                    Names.Add(element.ToString());
                }
            }

            if (Names.Count > 0)
            {
                foreach (string element in Names)
                {
                    sReg += element + ";";
                }
                rh.SaveSetting(@"Settings\Pfade", "TextFileList", sReg);
            }
        }

        private Collection<string[]> PrepareNameListCD(string sFile)
        {
            Collection<string[]> col = new Collection<string[]>();

            bool bName;
            string strName, strTmp;

            string[] arText;
            string[] arTitles;

            string sText = Action_Text_ReadFile(sFile);

            arText = sText.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            bName = IsName(arText);

            if (bName)
                strName = ExtractName(arText);
            else
                strName = "NoName";

            strName = strName.ToLowerInvariant();
            string strClip = "";

            for (int i = 0; i < arText.Length; i++)
            {
                strTmp = arText[i].Trim();

                if (strTmp.Length > 0 && strTmp.IndexOf(";") < 0 && strTmp.IndexOf("=") < 0 && strTmp.IndexOf("~") < 0)
                    strClip = string.Concat(strClip, strTmp, " # ");
            }

            strTmp = strClip;
            strTmp = strTmp.Replace((char)124, '#');
            strTmp = strTmp.Replace('*', '#');

            arTitles = strTmp.Split(new string[] { " # " }, StringSplitOptions.RemoveEmptyEntries);

            string strNr = "";
            string strTitel = "";
            int cnt = 1;

            foreach (string s in arTitles)
            {
                strNr = cnt.ToString("00");
                strTitel = strName + " - " + s.Trim();
                col.Add(new string[] { strNr, FormsUtilities.ToProperCase(strTitel) });
                cnt++;
            }

            return col;
        }

        private Collection<string[]> PrepareNameListCDAID_txt(string sFile)
        {
            Collection<string[]> col = new Collection<string[]>();

            string[] ar;
            string[] arTexte;
            string strTitel = "";

            string sText = Action_Text_ReadFile(sFile);

            ar = sText.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            short n = 0;
            string strNr = "";
            string strTmp = ar[1];

            arTexte = strTmp.Trim().Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            strTitel = arTexte[2] + " = " + arTexte[3];

            foreach (string Element in ar)
            {
                if (n > 0)
                {
                    strTmp = Element;
                    if (strTmp.Trim().Length > 0)
                    {
                        arTexte = strTmp.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                        if (arTexte[2].ToLower() == "various")
                            strTitel = arTexte[9].Replace("/", "-");
                        else
                            strTitel = arTexte[2] + " - " + arTexte[9];

                        strNr = n.ToString("00");
                        col.Add(new string[] { strNr, strTitel });
                    }
                }
                n++;
            }
            return col;
        }

        private Collection<string[]> PrepareNameListCDAID_xml(string sFile)
        {
            string artist;
            string cdtitel;
            string genre;
            string snr;
            string strack;
            string[] ar;

            List<string> tracklist = new List<string>();

            Collection<string[]> col = new Collection<string[]>();

            CDAID.XMLFile = sFile;
            CDAID.ReadDb();
            artist = CDAID.Artist;
            cdtitel = CDAID.CDTitel;
            genre = CDAID.Genre;
            tracklist = CDAID.TrackList;

            char[] charSeparators = new char[] { '/','-' };

            if (artist == "Various")
            {
                snr = "";
                strack = "";
                
                foreach (string track in tracklist)
                {
                    ar = track.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                    snr = ar[0];
                    strack = ar[1];
                    ar = strack.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
                    artist = ar[0];
                    col.Add(new string[] { snr, artist.Trim() + " - " + ar[1].Trim() });
                }
            }
            else
            {
                foreach (string track in tracklist)
                {
                    ar = track.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                    col.Add(new string[] { ar[0], artist + " - " + ar[1] });
                }
            }  
            return col;
        }

        private Collection<string[]> PrepareNameListDB(string sFile)
        {
            Collection<string[]> col = new Collection<string[]>();

            int ctr = 0;
            string[] ar;
            int intLmargin = 10;

            string sText = Action_Text_ReadFile(sFile);

            ar = sText.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var Element in ar)
            {
                if (Element.Trim() != "")
                {
                    string strTmp = Element;
                    if (strTmp.Trim().Length > 0 && strTmp.IndexOf("=") == -1 && strTmp.Substring(0, 1) != "~" && strTmp.IndexOf(";") == -1)
                    {
                        ++ctr;
                        string strNr = ctr.ToString("00");
                        strTmp = strTmp.TrimEnd();
                        strTmp = strTmp.Substring(intLmargin - 1);
                        int l = strTmp.IndexOf("  ");

                        string strTitel = strTmp.Substring(0, l);
                        string strName = strTmp.Substring(l, strTmp.Length - l).TrimStart();

                        col.Add(new String[] { strNr, strName + " - " + strTitel });
                    }
                }
            }

            return col;
        }

        private void Action_ComboNames()
        {
            string strFileName = "";

            if (ComboBoxTitles.Text.Contains("***"))
            {
                strFileName = "c:\\temp\\paste.txt";
            }
            else
            {
                strFileName = ComboBoxTitles.Text;
            }

            Action_Text_ReadFile(strFileName);
        }

        private void Action_Text_Wechseln()
		{
			foreach (ListViewItem itmX in listViewNames.Items)
			{
				if (itmX.SubItems[1].Text.IndexOf(" - ") > 0)
				{
					string [] ar = itmX.SubItems[1].Text.Split(new string[] { " - " }, StringSplitOptions.None);
					string strNeu = ar[1].Trim() + " - " + ar[0].Trim();
					itmX.SubItems[1].Text = strNeu;
				}
				else
				{
					itmX.ForeColor = System.Drawing.Color.Red;
					itmX.SubItems[1].ForeColor = System.Drawing.Color.Red;
				}
			}
		}

        public void Action_Text_Paste()
        {
            if (Clipboard.ContainsFileDropList())
            {
                StringCollection FileCol = new StringCollection();
                FileCol = Clipboard.GetFileDropList();

                MessageBox.Show("First File=" + FileCol[0], "Warnung");
            }

            string strClipText = "";

            if (!FormsUtilities.IsMenuChecked(this.menuMainTextformat))
            {
                MessageBox.Show("Es muss ein Textformat ausgewählt werden!", "Renamer - Text_Paste");
                return;
            }

            if (Clipboard.ContainsText())
            {
                strClipText = Clipboard.GetText();
            }

            if (strClipText.Trim() == "")
            {
                MessageBox.Show("Clipboard is empty!", "Action_Text_Paste");
                return;
            }

            PrepareNameList(strClipText);

            ComboBoxTitles.Text = "*** Pasted List ***";

            rh.SaveSetting("Settings\\Pfade", "LastTextFile", "Clipboard");

            string fileName = String.Concat(Path.GetTempPath(),_logFile);

            StreamWriter fs = new StreamWriter(fileName, true);

            fs.WriteLine(DateTime.UtcNow);
            fs.Write("action_text_past");
            fs.WriteLine("Filelist");

            foreach (ListViewItem itmX in this.listViewNames.Items)
            {
                fs.WriteLine(itmX.SubItems[1].Text);
            }

            fs.WriteLine("================================================================================");
            fs.Close();
        }

        public void Action_Text_Edit()
        {
            ListViewItem itmX = new ListViewItem();

            if (this.listViewNames.Items.Count == 0)
                return;

            if (this.listViewNames.SelectedItems.Count > 0)
            {
                itmX = this.listViewNames.FocusedItem;
                if (itmX.SubItems.Count == 0)
                    return;

                string strEingabe = "";

                int px = this.Left + itmX.Position.X;
                int py = this.Top + itmX.Position.Y;

                strEingabe = Interaction.InputBox("Listeintrag Editieren", "Edit", itmX.SubItems[1].Text, px, py);

                if (strEingabe != "")
                {
                    foreach (ListViewItem itmY in this.listViewNames.Items)
                        itmY.BackColor = Color.Black; // Color.White;

                    itmX.SubItems[1].Text = strEingabe;
                    string strVorbidden = "";
                    bool bForbChars = false;

                    foreach (ListViewItem itmY in this.listViewNames.Items)
                    {
                        strVorbidden = CheckForbiddenCharacters(itmY.SubItems[1].Text);
                        if (strVorbidden != "")
                        {
                            itmY.BackColor = Color.LightSalmon;
                            bForbChars = true;
                        }
                    }

                    if (bForbChars == true)
                    {
                        MessageBox.Show("Beim Einlesen wurden unerlaubte Zeichen gefunden!", "Renamer - Text" + strVorbidden);
                    }


                }
            }
        }

        private void Action_File_TouchFiles()
        {
            string strFolder = this.ComboBoxFiles.Text;
            string oldDate, strFileName;

            oldDate = this.listViewFiles.Items[0].SubItems[3].Text;
            string sT1 = oldDate.Substring(0, oldDate.Length - 2);

            short counter = 1;
            foreach (ListViewItem li in listViewFiles.Items)
            {
                string sT2 = (sT1 + counter.ToString("00"));
                DateTime dnewDate = DateTime.Parse(sT2);
                strFileName = strFolder + "\\" + li.SubItems[1].Text + li.SubItems[2].Text;
                bool retval = fh.TouchFile(strFileName, sT2);
                if (retval == true)
                {
                    li.SubItems[3].Text = sT2;
                }
                counter++;
            }
        }

        private void Action_File_RenameDo()
        {
            string strNr = "", strDate = "", strTitel = "", keyPath = "";
            int n = 0, intSec = 0, intMin = 0, intStd = 0;
            bool renameSuccess = false;

            ///
            /// Save Settings for the Undo Function
            ///
            /// delete old settings
            ///

            string path = ComboBoxFiles.Text;

            if(!Directory.Exists(path))
            {
                MessageBox.Show("Source folder not found!","Renamer - File");
                return;
            }

            keyPath = _keyRoot + "\\Settings\\Names";
            rh.DeleteSubKeyTree(RegistryRootKeys.HKEY_CURRENT_USER, keyPath);

            keyPath = _keyRoot + "\\Settings\\Files";
            rh.DeleteSubKeyTree(RegistryRootKeys.HKEY_CURRENT_USER, keyPath);
            
            ///
            ///	write new setting
            ///
            foreach (ListViewItem itmNames in this.listViewNames.Items)
            {
                strNr = String.Format(itmNames.Text,"00");
                strTitel = itmNames.SubItems[1].Text;
                rh.SaveSetting("Settings\\Names", strNr, strTitel);
            }

            strNr = "";
            strTitel = "";

            foreach (ListViewItem itmFiles in this.listViewFiles.Items)
            {
                strNr = String.Format(itmFiles.Text,"00");
                strTitel = itmFiles.SubItems[1].Text;
                rh.SaveSetting("Settings\\Files", strNr, strTitel);
            }

            //
            // Log schreiben und file umbenennen
            //

            string renameID = "";
            keyPath = _keyRoot + "\\Settings";

            try
            {
                var retVal = rh.GetSetting("Settings", "renameID", "-1");
                newID = Convert.ToInt64(retVal);
                newID += 1;
                renameID = newID.ToString("00000");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return;
            }

			string sT1, strPfad, strFileNameOld, strExt;
			string oldDate;
			string fileName = String.Concat(gh.ValidatePath(Application.StartupPath), _logFile);

            DateTime timestamp = new DateTime();
            timestamp = DateTime.UtcNow;

            strDate = timestamp.ToShortDateString();
            intStd = timestamp.Hour;
            intMin = timestamp.Minute;

            // olddate is not avaiable on redo action
			oldDate = this.listViewFiles.Items[0].SubItems[3].Text;
            if (!String.IsNullOrEmpty(oldDate))
            {
                strDate = oldDate.Substring(0, 10);
                sT1 = oldDate.Substring(11, 2);
                intStd = Convert.ToInt32(oldDate.Substring(11, 2));
                sT1 = oldDate.Substring(14, 2);
                intMin = Convert.ToInt32(oldDate.Substring(14, 2));
            }
            
            strPfad = gh.ValidatePath(ComboBoxFiles.Text);

            ArrayList logEntries = new ArrayList();

            logEntries.Add(String.Format("[{0}] [{1}] source {2}", timestamp, renameID, strPfad));
            
            strTitel = "";
			foreach ( ListViewItem itmNames in this.listViewNames.Items)
			{
				n = itmNames.Index;
				ListViewItem itmFiles = listViewFiles.Items[n];

				strExt = "." + itmFiles.SubItems[2].Text.ToLower();
				strTitel = itmNames.SubItems[1].Text;
				strFileNameOld = itmFiles.SubItems[1].Text + strExt;

				intSec += 1;
				if (intSec > 59)
				{
					intSec = 0;	intMin += 1;
					if (intMin > 59) intMin = 0;
				}

				string sT2 = (strDate + " " + intStd.ToString("00") + ":" + intMin.ToString("00") + ":" + intSec.ToString("00"));
				DateTime dT2 = DateTime.Parse(sT2);

                logEntries.Add(String.Format("[{0}] [{1}] {2}", dT2, renameID, ("rename \"" + strFileNameOld + "\" ==> \"" + strTitel + strExt + "\"")));

				// alten titel umbenennen
				try
				{
					File.Move(strPfad + strFileNameOld, strPfad + strTitel + strExt);
					bool retval = fh.TouchFile(strPfad + strTitel + strExt, sT2);
					itmFiles.SubItems[1].Text = strTitel;
                    itmFiles.BackColor = Color.FromArgb(185,250,185);
                    renameSuccess = true;
				}
				catch
				{
					itmFiles.BackColor = Color.MistyRose;
                    itmFiles.ForeColor = Color.DarkBlue;
                    renameSuccess = false;
				}
			}

            if (renameSuccess == true)
            {
                StreamWriter writer = new StreamWriter(fileName, true);
                foreach (string line in logEntries)
                {
                    writer.WriteLine(line);
                }
                writer.Close();

                newID = Convert.ToInt64(renameID);

                keyPath = _keyRoot + "\\Settings";

                rh.SetValue(RegistryRootKeys.HKEY_CURRENT_USER, keyPath, "renameID", newID, true, RegistryValueKind.DWord);

                for (n = 0; n < listViewFiles.Columns.Count; ++n)
                {
                    listViewFiles.AutoResizeColumn(n, ColumnHeaderAutoResizeStyle.ColumnContent);
                }

                SplitContainer1.BackColor = Color.Salmon;
            }
            else
            {
                MessageBox.Show("File rename failed!", "File_Rename");
            }
        }

        private void Action_File_RenamePrepUndo()
        {
            StreamReader reader;
            List<string> logdates = new List<string>();
            List<Tracks> listNames = new List<Tracks>();
            List<Tracks> listFiles = new List<Tracks>();

            listViewNames.Items.Clear();
            listViewFiles.Items.Clear();
            SplitContainer1.BackColor = Color.Silver;
            ComboBoxFiles.Text = "";
            ComboBoxTitles.Text = "";

            Choose choose = new Choose();

            int x = Location.X + 60;
            int y = Location.Y + 120;
            choose.Location = new Point(x, y);
            choose.ShowDialog(this);
            LogID = choose.LogID;

            choose.Close();
            choose.Dispose();

            if (LogID <= 0) return;

            string renameID = LogID.ToString("00000");
            string sourcePath = "";
            string line;
            reader = new StreamReader(_logFile);

            while (!reader.EndOfStream)
            {
                line = reader.ReadLine();

                if (line.IndexOf(renameID) > 0 && line.IndexOf("source") > 0)
                {
                    int i = line.IndexOf("source") + "source".Length + 1;
                    sourcePath = line.Substring(i, line.Length - i);
                }

                if (line.IndexOf(renameID) > 0 && line.IndexOf("rename") > 0)
                {
                    string[] ar = line.Split('"');
                    string nam = ar[1].Replace("\"", "");
                    listNames.Add(new Tracks { name = nam.Replace(".mp3", "") });

                }

                if (line.IndexOf(renameID) > 0 && line.IndexOf("==>") > 0)
                {
                    string[] ar = line.Split('"');
                    string fil = ar[3].Replace("\"", "");
                    listFiles.Add(new Tracks { name = fil.Replace(".mp3", ""), extension = "mp3" });
                }
            }

            reader.Close();
            reader.Dispose();

            ComboBoxFiles.Text = sourcePath;

            FillListView(listViewNames, listNames);
            FillListView(listViewFiles, listFiles);

        }

        public void Action_File_Paste()
        {
            if (Clipboard.ContainsFileDropList())
            {
                StringCollection FileCol = new StringCollection();
                FileCol = Clipboard.GetFileDropList();

                DirectoryInfo di;
                di = new DirectoryInfo(FileCol[0]);
                string strFolder;

                if ((di.Attributes & FileAttributes.Directory) > 0)
                    strFolder = di.FullName;
                else
                    strFolder = di.Parent.FullName;

                ComboBoxFiles.Text = strFolder;
            }
            else if(Clipboard.ContainsText())
            {
                ComboBoxFiles.Text = Clipboard.GetText();
            }

        }

        public void Action_File_SelectAll()
        {
            if (listViewFiles.Items.Count > 0)
            {
                foreach (ListViewItem xItm in listViewFiles.Items)
                {
                    xItm.Selected = true;
                }
            }
        }

        public int Action_File_FindItem(string URL)
        {
            int item = -1;

            foreach (ListViewItem xItm in listViewFiles.Items)
            {
                string strItem = xItm.SubItems[1].Text;
                item = xItm.Index;
                if (URL.Contains(strItem)) break;
            }

            return item;
        }

        private void Action_CopyToClip(ListView LV, short Index)
        {
            string strList = "";
            string strLine = "";
            string[] ar;

            switch (Index)
            {
                case 0: //  ' as CD

                    ar = LV.Items[1].SubItems[1].Text.Split(new string[] { " - " }, StringSplitOptions.None);
                    strList = "".PadLeft(11) + ar[0].ToUpper() + " = " + "\r\n\n";
                    strLine = "".PadLeft(4);

                    foreach (ListViewItem xItm in LV.Items)
                    {
                        ar = xItm.SubItems[1].Text.Split(new string[] { " - " }, StringSplitOptions.None);
                        if (strLine.Length + ar[1].Length < 76)
                        {
                            strLine += ar[1] + " * ";
                        }
                        else
                        {
                            strLine = strLine.Substring(1, strLine.Length - 4);
                            strList += strLine + "\r\n";
                            strLine = "".PadLeft(4) + ar[1] + " * ";
                        }
                    }
                    break;

                case 1: //  ' as DB
                    string strTmp = "";

                    foreach (ListViewItem xItm in LV.Items)
                    {
                        ar = xItm.SubItems[1].Text.Split(new string[] { " - " }, StringSplitOptions.None);
                        if (ar[1].ToString().Length < 46)
                            strTmp = ar[1];
                        else
                            strTmp = ar[1].ToString().Substring(0, 46);

                        strLine = xItm.Text + "".PadLeft(7) + strTmp + "".PadLeft(50 - strTmp.Length) + ar[0];
                        strList += strLine + "\r\n";
                    }
                    break;

                case 2: //  ' as is
                    foreach (ListViewItem xItm in LV.Items)
                    {
                        strList += xItm.SubItems[1].Text + "\r\n";
                    }
                    break;
            }

            Clipboard.Clear();
            Clipboard.SetText(strList);
        }

        internal void Action_MoveUp(ListView lv)
        {
            int n = lv.SelectedItems[0].Index;           
            if (n == 0)
                return;
                
            ListViewItem xItm1 = lv.Items[lv.SelectedItems[0].Index];

            xItm1.Remove();
            lv.Items.Insert(n -1, xItm1);

            int i = 0;
            foreach (ListViewItem li in lv.Items)
            {
                i++;
                li.Text = i.ToString("00");
            }
        }

        internal void Action_MoveDown(ListView lv)
        {
            int n = lv.SelectedItems[0].Index;
            if (n == lv.Items.Count-1)
                return;

            ListViewItem xItm1 = lv.Items[lv.SelectedItems[0].Index];

            xItm1.Remove();
            lv.Items.Insert(n + 1, xItm1);

            int i = 0;
            foreach (ListViewItem li in lv.Items)
            {
                i++;
                li.Text = i.ToString("00");
            }
        }

        internal void Action_Cut(ListView LV)
        {
            if (LV.Items.Count == 0)
                return;

            int sel = LV.SelectedItems[0].Index;
            foreach (ListViewItem li in LV.SelectedItems)
            {
                li.Remove();
            }

            if (LV.Items.Count == 0)
                return;

            if(sel < LV.Items.Count && LV.Items.Count >0)
                LV.Items[sel].Selected = true;
            else
                LV.Items[LV.Items.Count-1].Selected = true;

            int i = 0;
            foreach (ListViewItem li in LV.Items)
            {
                i++;
                li.Text = i.ToString("00");

            }

            this.EnableRename();
        }

        internal string Action_Text_ReadFile(string fileName)
        {
            string strText = "";
			if (File.Exists(fileName))
			{
				StreamReader reader = new StreamReader(fileName);
				strText = reader.ReadToEnd();
				reader.Close();
                reader.Dispose();
			}
			else
			{
				if (fileName != "")
				{
					strText = fileName;
				}
			}
			return strText;
        }

        /// <summary>
        /// fill listviewfiles with all files in folder shown in comboboxfiles
        /// </summary>
        /// <param name="strFolder"></param>
        private void PrepareFileList(string strFolder)
        {
            string strNr = "";
            string strFileName = "";
            string strFileExtension = "";
            string strExtension = "";
            string strFileDate = "";
            string strFileSize = "";

            List<Tracks> tracks = new List<Tracks>();

            // listViewFiles.ForeColor = Color.White;
            listViewFiles.Font = new Font(listViewFiles.Font, FontStyle.Regular);

            if (strFolder.Trim().Length == 0) return;

            DirectoryInfo di;
            di = new DirectoryInfo(strFolder);

            if (!di.Exists) return;

            var fileinfos = di.GetFiles("*.mp3").Concat(di.GetFiles("*.wav"));

            int cnt = 1;

            foreach (FileInfo fi in fileinfos)
            {
                if (fi.Exists)
                {
                    strExtension = fi.Extension.ToLower();
                    strNr = cnt.ToString("00");
                    strFileName = fi.Name.Replace(strExtension, "");
                    strFileDate = fi.CreationTime.ToUniversalTime().ToString();
                    strFileSize = (fi.Length / 1024).ToString("###,### KB");
                    tracks.Add(new Tracks { nr = strNr, name = strFileName, extension = strExtension.Replace(".", ""),
                                        fileDate = strFileDate, fileSize = strFileSize });

                    cnt++;
                }
            }
           
            rh.SaveSetting("Settings\\Pfade", "LastMp3Dir", ComboBoxFiles.Text);
            // FillListView(listViewFiles, col);
            FillListView(listViewFiles, tracks);

            this.statusStripFilecount.Text = listViewFiles.Items.Count.ToString();
        }

        private void PrepareNameList(string sFileName)
        {
			bool readerror = false;

            if (sFileName.Trim().Length == 0)
                return;

            int n = FormsUtilities.GetCheckedMenuItem(menuMainTextformat);

            List<Tracks> listTracks = new List<Tracks>();

            Collection<string[]> col = new Collection<string[]>();

            switch (n)
            {
                case 0:
                    //Menü Auto
                    Console.WriteLine("Case 0");
                    break;
                case 1:
					// Menü ASIS
                    col = PrepareNameListASIS(sFileName);
                    // Debug.Print("Case 1");
                    break;
                case 2:
					// Menü CD
					Console.WriteLine("Case 2");
                    col = PrepareNameListCD(sFileName);
                    break;
                case 3:
					// Menü CDAid
                    if (sFileName.IndexOf(".xml") >0)
                        col = PrepareNameListCDAID_xml(sFileName);
                    if (sFileName.IndexOf(".txt") >0)
                        col = PrepareNameListCDAID_txt(sFileName);
                    Console.WriteLine("Case 3");
                    break;
                case 4:
					// Menü DB
					if (sFileName.IndexOf((char)4) > 0)
					{
						MessageBox.Show("seams this is a CD formated text!", "Renamer - File");
						readerror = true;
						return;
					}
                    col = PrepareNameListDB(sFileName);
                    Console.WriteLine("Case 4");
                    break;
            }

			if (readerror == true)
				return;

            rh.SaveSetting("Settings\\Pfade", "LastTextFile", ComboBoxTitles.Text);
            FillListView(listViewNames, listTracks);

            this.statusStripNamecount.Text = listViewNames.Items.Count.ToString();

            string strVorbidden = "";
            bool bForbChars = false;

            foreach (ListViewItem itmX in this.listViewNames.Items)
            {
                strVorbidden = CheckForbiddenCharacters(itmX.SubItems[1].Text);

                if (strVorbidden != "")
                {
                    itmX.BackColor = Color.LightSalmon;
                    bForbChars = true;
                }
            }

            if (bForbChars == true)
            {
                MessageBox.Show("Beim Einlesen wurden unerlaubte Zeichen gefunden!", "Unerlaubtes Zeichen " + strVorbidden);
            }

            FormsUtilities.UncheckMenu(menuMainTextformat);

            listViewNames.Refresh();
        }

        string CheckForbiddenCharacters(string strText)
        {
            int n;
            string c = "";

            string strForbChars = (string)rh.GetSetting("Settings", "ForbChars", "\\/:?");

            for (n = 0; n < strForbChars.Length; ++n)
            {
                char t = strForbChars[n];

                if (strText.IndexOf(t) > 0)
                {
                    c = t.ToString();
                }
            }

            return c;
        }
        
        internal void FillCombo(ComboBox combo, string key)
        {
            string[] ar;

            combo.Items.Clear();
            combo.Items.Add("");
            var regval = (string)rh.GetSetting(@"Settings\Pfade", key, "");
            if (String.IsNullOrEmpty(regval)) return;
            ar = regval.Split(new string[] {";"}, StringSplitOptions.RemoveEmptyEntries);

            foreach (string element in ar)
            {
                if (element.Trim() != "")
                    Add2combo(combo, element);
            }
            combo.SelectedItem = 0;
        }

        private void FillListView(ListView LV, List<Tracks> listTracks)
        {
            ListViewItem itmX;
            int imageIndex = 0;

            if (LV.Name == "listViewNames")
                imageIndex = 0;

            if (LV.Name == "listViewFiles")
                imageIndex = 1;

            LV.Items.Clear();
            LV.BeginUpdate();

            int m = 1;

            foreach (Tracks track in listTracks)
            {
                itmX = LV.Items.Add(m.ToString("00"));
                itmX.ImageIndex = 0;
                itmX.SubItems.Add(track.name);
                itmX.SubItems.Add(track.extension);
                itmX.SubItems.Add(track.fileDate);
                itmX.SubItems.Add(track.fileSize);
                itmX.ImageIndex = imageIndex;
                m++;
            }

            LV.EndUpdate();

            for (int n = 0; n < LV.Columns.Count; ++n)
            {
                LV.AutoResizeColumn(n, ColumnHeaderAutoResizeStyle.ColumnContent);
            }

            EnableRename();
        }

        public void InitListViews()
        {

            // Create the ListView image lists.
            ImageList smallImageList = new ImageList();
            smallImageList.ImageSize = new Size(16, 16);
            ImageList largeImageList = new ImageList();
            largeImageList.ImageSize = new Size(32, 32);

            // Load the document icon.
            Icon doc = Properties.Resources.Dokument;
            Icon note = Properties.Resources.Note;

            smallImageList.Images.Add(doc);
            largeImageList.Images.Add(doc);
            this.listViewNames.SmallImageList = smallImageList;
            
            smallImageList.Images.Add(note);
            largeImageList.Images.Add(note);
            this.listViewFiles.SmallImageList = smallImageList;
        }

        private void EnableRename()
        {
            if (menuMainExtrasDirectMode.Checked)
            {
                if (listViewFiles.Items.Count > 0 && listViewNames.Items.Count > 0)
                {
                    menuMainFileRename.Enabled = true;
                    SplitContainer1.BackColor = Color.LightGreen;
                }
                else
                {
                    menuMainFileRename.Enabled = false;
                    SplitContainer1.BackColor = Color.Silver;
                }
            }
            else
            {
                if (listViewFiles.Items.Count > 0 && listViewFiles.Items.Count == listViewNames.Items.Count)
                {
                    menuMainFileRename.Enabled = true;
                    SplitContainer1.BackColor = Color.LightGreen;

                }
                else
                {
                    menuMainFileRename.Enabled = false;
                    SplitContainer1.BackColor = Color.Silver;
                }
            }
        }

        internal string ExtractName(string[] ar)
        {
            string strName = "";
            string strZeile = "";
            string strTmp = "";

            for (short i = 0; i < ar.Length; i++)
            {
                if (ar[i].IndexOf("=") > 0)
                {
                    strZeile = ar[i];
                    break;
                }
            }

            int l = strZeile.Length;
            int n = strZeile.IndexOf("=");

            if (n > 1)
            {
                strTmp = strZeile.Substring(0, n - 1).Trim();
                strTmp = strTmp.Substring(10, strTmp.Length - 10).Trim();
                strName = strTmp;
            }

            return strName;
        }

        internal bool IsName(string[] ar)
        {
            for (short i = 0; i < ar.Length; i++)
            {
                if (ar[i].IndexOf("=") > 0)
                {
                    return true;
                }
            }
            return false;
        }

        public bool Add2combo(System.Windows.Forms.ComboBox cmb, string sText)
        {
            if (sText.Trim() == string.Empty)
                return true;

            sText = sText.ToLower();

            Collection<string> col = new Collection<string>();

            foreach (string element in cmb.Items)
            {
                if (element.Trim() != string.Empty)
                    col.Add(element);
            }

            if ( !col.Contains(sText))
                cmb.Items.Add(sText);

            return true;
        }

        public bool Add2ListView(System.Windows.Forms.ListView lv, string sText)
        {
            if (sText.Trim() == string.Empty)
                return true;

            sText = sText.ToLower();

            Collection<string> col = new Collection<string>();

            foreach (string element in lv.Items)
            {
                if (element.Trim() != string.Empty)
                    col.Add(element);
            }

            if (!col.Contains(sText))
                lv.Items.Add(sText);

            return true;
        }

        private Collection<string[]> PrepareNameListASIS(string sFile)
        {
            string sText = Action_Text_ReadFile(sFile);

            Collection<string[]> col = new Collection<string[]>();

            string strNr = "";
            int cnt = 0;

			string[] ar = sText.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string n in ar)
            {
                if (n.Trim() != "")
                {
                    ++cnt;
                    strNr = cnt.ToString("00");
                    col.Add(new string[] { strNr, n });
                }

            }
            return col;
        }

        private void Player_PlayStateChange(int NewState)
        {
            //wmppsUndefined = 0,
            //wmppsStopped = 1,
            //wmppsPaused = 2,
            //wmppsPlaying = 3,
            //wmppsScanForward = 4,
            //wmppsScanReverse = 5,
            //wmppsBuffering = 6,
            //wmppsWaiting = 7,
            //wmppsMediaEnded = 8,
            //wmppsTransitioning = 9,
            //wmppsReady = 10,
            //wmppsReconnecting = 11,
            //wmppsLast = 12

            if (NewState == (int)WMPLib.WMPPlayState.wmppsMediaEnded)
            {
                if (this.checkBoxPlayOption.Checked)
                {

                }
            }
        }

        private void Player_MediaError(object pMediaObject)
        {
            MessageBox.Show("Cannot play media file.");
        }

        #endregion public Methodes

        private void contextMenuFilesToProperCase_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem li in listViewFiles.SelectedItems)
            {
                string strText = FormsUtilities.ToProperCase(li.SubItems[1].Text);
                li.SubItems[1].Text = strText;
            }
        }
    }
    #endregion class Renamer

    #region class ListViewItemComparer
    // Implements the manual sorting of items by columns.
    public class ListViewItemComparer : IComparer
	{
		private int currentColumn;
		private SortOrder sortorder;
		
        public SortOrder SortOrder
		{
			get { return this.sortorder; }
			set { this.sortorder = value; }
		}

		public int CurrentColumn
		{
			get { return this.currentColumn; }
			set { this.currentColumn = value; }
        }

        public ListViewItemComparer()
		{
			currentColumn = 0;
		}

		public ListViewItemComparer(int column)
		{
			currentColumn = column;
		}

		public int Compare(object x, object y)
		{
			return String.Compare(((ListViewItem) x).SubItems[currentColumn].Text, ((ListViewItem) y).SubItems[currentColumn].Text);
        }
    }
    #endregion class ListViewItemComparer

}
