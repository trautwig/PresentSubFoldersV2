using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;

namespace PresentSubfolders
{
    class SearchLeader
    {
        BackgroundWorker _worker;
        public SearchLeader(BackgroundWorker worker)
        {
            _worker = worker;
        }

        public void doSearch()
        {
            //get the top level of subfolders
            List<string> allDirectories;
            Program.maxLevel = 0;
            string currentFolderName;
            int slashPosition;
            SubFolder currentFolder;
            allDirectories = new List<string>();
            string[] rootSubdirectoryEntries = Directory.GetDirectories(Program.searchString, "*", SearchOption.TopDirectoryOnly);

            //some correction for recording depth levels
            Program.slashCorrection = Program.searchString.Count(s => s == '\\');
            if (Program.slashCorrection == 1)
            {//if you're in the top level of a drive, there is no need to correct
                Program.slashCorrection = 0;
            }
            //pull back just the top level
            Program.centralFiles = new TopLevelFolder("T", 1, "");

            //now pull back all the next top-level subfolders in the ones you've already retrieved
            double max = rootSubdirectoryEntries.Length;//this and the following variable will be used for reporting progress of work back to UI
            double report;
            int i = 0; //assigning i here so that it can be used oustide the for loop
                       //            folderRecordingProgressBar.Minimum = 0;
                       //            folderRecordingProgressBar.Maximum = rootSubdirectoryEntries.Length - 2;
                       //            folderRecordingProgressBar.Step = 1;
            for (i = 0; i < max; i++)
            //possibly add progress bar here
            {
                if (rootSubdirectoryEntries[i].IndexOf("Admin") < 0 &&
                    rootSubdirectoryEntries[i].IndexOf("Back up (Restored)") < 0
                    && rootSubdirectoryEntries[i].IndexOf("bu of electronic filing") < 0)
                //don't bother with these directories
                {
                    slashPosition = rootSubdirectoryEntries[i].LastIndexOf("\\");
                    currentFolderName = rootSubdirectoryEntries[i].Substring(slashPosition + 1);
                    if (currentFolderName.IndexOf(",") != -1)
                    {
                        currentFolderName.Replace(",", "','");
                    }
                    currentFolder = Program.centralFiles.createDeeperFolder(currentFolderName, 1, "");
                    //since you're just working in top level folders,
                    //call populateSubFolders for each top level folder; that function will 
                    //record all the subfolders in the top level folder
                    Program.centralFiles.populateSubFolders(currentFolder, rootSubdirectoryEntries[i]);
                }
                report = System.Convert.ToDouble(i) / max;//calculate progress of work
                report = report * 100;
                _worker.ReportProgress(System.Convert.ToInt32(report));//report progress to UI
            }
            if (i == rootSubdirectoryEntries.Length && Form1.currentSubFolderDone)
            {
                Form1.topLevelFoldersDone = true;
            }
        }


    }
}
