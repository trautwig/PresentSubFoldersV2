using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace PresentSubfolders
{
    public partial class Form1 : Form
    {
        public static bool currentSubFolderDone;
        public static bool topLevelFoldersDone;
        public static string excelFileLinkText;
        public static string wordFileLinkText;

        string outputMode;// this will track whether the list of folders is being made in word or excel

        public Form1()
        {
            InitializeComponent();
            bgWorkerFolderSearch.WorkerReportsProgress = true;
            bgWorkerFolderSearch.WorkerSupportsCancellation = true;
            bgWorkerResultsToExcelOrWord.WorkerReportsProgress = true;
            bgWorkerResultsToExcelOrWord.WorkerSupportsCancellation = true;
            badResultsPanel.Visible = false;
            workingOnFilesPanel.Visible = false;
            levelDisplayGroup.Visible = false;
            excelResultsDonePanel.Visible = false;
            workingOnWordPanel.Visible = false;
            wordResultsDonePanel.Visible = false;
            //!!! maybe move excel and word objects into global scope in this form
            //so that they can be killed on quit
        }

        /// <summary>
        /// Kicks off the search for subfolders. It relies on SubFolder.TopLevelFolder.populateSubFolders() to do most of the work.
        /// </summary>
        private void startSearch()
        {
            //set display items for searching mode
            badResultsPanel.Visible = false;
            workingOnWordPanel.Visible = false;
            currentSubFolderDone = false;
            topLevelFoldersDone = false;
            folderRecordingProgressBar.Value = 0;
            folderRecordingProgressBar.Maximum = 100;
            workingOnFilesPanel.Visible = true;
            workingOnFilesPanel.Update();
            if(levelDisplayGroup.Visible)
            {
                levelDisplayGroup.Visible = false;
                //reset the text about the depth level
                depthLevelText.Text = "The folders found went to a depth of";
            }
            if (excelResultsDonePanel.Visible)
            {
                excelResultsDonePanel.Visible = false;
            }
            if (wordResultsDonePanel.Visible)
            {
                wordResultsDonePanel.Visible = false;
            }
            bgWorkerFolderSearch.RunWorkerAsync();//that creates a SerarchLeader, which is where the search work is done. 
        }

        /// <summary>
        /// Allows user to select the root folder from which they start recording subfolders.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectFolders(object sender, EventArgs e)
        {
            excelFileLink.Text = "";
            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                Program.searchString = Path.GetDirectoryName(folderBrowser.FileName);
                startSearch();
            }
        }

        /// <summary>
        /// Event handler for clicking on the link to the created excel file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(((System.Windows.Forms.LinkLabel)sender).Text);
        }

        private void depthLevelsToDisplay_ValueChanged(object sender, EventArgs e)
        {

        }

        private void createExcel_Click(object sender, EventArgs e)
        {//user wants results in excel, make that happen
            outputMode = "excel";
            if (excelResultsDonePanel.Visible)
            {
                excelResultsDonePanel.Visible = false;
            }
            if (wordResultsDonePanel.Visible)
            {
                wordResultsDonePanel.Visible = false;
            }
            bgWorkerResultsToExcelOrWord.RunWorkerAsync();
        }

        private void createWord_Click(object sender, EventArgs e)
        {//user wants results in word, make that happen
            outputMode = "word";
            if (excelResultsDonePanel.Visible)
            {
                excelResultsDonePanel.Visible = false;
            }
            if (wordResultsDonePanel.Visible)
            {
                wordResultsDonePanel.Visible = false;
            }
            workingOnWordPanel.Visible = true;
            workingOnWordProgressBar.Value = 0;
            workingOnWordProgressBar.Step = 1;
            workingOnWordProgressBar.Maximum = 100;
            workingOnWordPanel.Update();

            bgWorkerResultsToExcelOrWord.RunWorkerAsync();
            //recordResultsToWord();
        }

//event hanlders for bgWorkerFolderSearch
        private void bgWorkerFolderSearch_DoWork(object sender, DoWorkEventArgs e)
        {
            //this will fire the search worker and pass sender to it (to allow report back to this UI)
            BackgroundWorker worker = sender as BackgroundWorker;
            SearchLeader searchLeader = new SearchLeader(worker);
            searchLeader.doSearch();
        }

        private void bgWorkerFolderSearch_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            folderRecordingProgressBar.Value = e.ProgressPercentage;
        }

        private void bgWorkerFolderSearch_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //folder search is done, change display stuff
            //display changes to record the end of searching mode
            workingOnFilesPanel.Visible = false;
            workingOnFilesPanel.Update();
            //display a max depth, 
            if (topLevelFoldersDone)
            {
                //prompt user to select a depth to display to
                badResultsPanel.Visible = false;
                depthLevelText.Text = depthLevelText.Text + " " + Program.maxLevel.ToString();
                depthLevelsToDisplay.Maximum = Program.maxLevel;
                levelDisplayGroup.Visible = true;
            }
            else
            {
                levelDisplayGroup.Visible = false;
                badResultsPanel.Visible = true;
            }
        }

//event handlers for bgWorkerResultstoExcelOrWord
        private void bgWorkerResultsToExcelOrWord_DoWork(object sender, DoWorkEventArgs e)
        {   //this will fire the results recorder and pass sender to it (to allow report back to this UI)
            BackgroundWorker worker = sender as BackgroundWorker;
            ResultsRecorder resultsRecorder = new ResultsRecorder(worker);
            resultsRecorder.depthLevelsToDisplay = System.Convert.ToInt32(depthLevelsToDisplay.Value);
            if (outputMode == "word")
            { 
            }
            resultsRecorder.recordResultsToExcelOrWord(outputMode);
        }

        private void bgWorkerResultsToExcelOrWord_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (outputMode == "word")
            {
                workingOnWordProgressBar.Value = e.ProgressPercentage;
            }//right now, there is only progress reported back if the recording is in word (because excel runs quickly)
        }

        private void bgWorkerResultsToExcelOrWord_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {//update UI based on completion of recording folder info to word or excel
            if (outputMode == "excel")
            { 
                excelFileLink.Text = excelFileLinkText;
                excelResultsDonePanel.Visible = true;
                excelFileLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            }
            else if (outputMode == "word")
            {
                workingOnWordPanel.Visible = false;
                wordResultsDonePanel.Visible = true;
                wordFileLink.Text = wordFileLinkText;
                wordFileLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            }
        }
    }
}
