using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PresentSubfolders
{
    public class SubFolder
    {
        protected string _name;
        protected int _level;
            //for folder "T:\\BU - BUDGET AND FALL ECONOMIC STATEMENTS\\2021\\01 - Binder & Lock Up"
            //_level is 3
        protected string _parentFolder;
        
        public Queue <SubFolder> subFolders = new Queue<SubFolder>();
        
        public SubFolder(string sfName, int sfLevel, string sfParentFolder)
        {
            _name = sfName;
            _level = sfLevel;
            _parentFolder = sfParentFolder;
            if (_level > Program.maxLevel )
            {
                Program.maxLevel = _level;//make sure we record the deepest subFolder 
            }
        }

        public string sfName
        {
            get { return _name; }
        }

        public int sfLevel
        {
            get { return _level; }
        }

        public string sfParentFolder
        {
            get { return _parentFolder; }
        }


        public SubFolder createDeeperFolder(string nlName, int nlLevel, string nlParentFolder)
        {
            SubFolder sf = new SubFolder(nlName, nlLevel, nlParentFolder);
            subFolders.Enqueue(sf);
            return sf;
        }



    }

    public class TopLevelFolder : SubFolder
    {
        public TopLevelFolder(string sfName, int sfLevel, string sfParentFolder) : base(sfName, sfLevel, sfParentFolder)
        {

        }

        /// <summary>
        /// This function will find the sub-folders of inputFolder and put them into inputFolder's subFolders property.
        /// This is a recursive function: if inputFolder itself has subfolders, each of those subFolders will 
        /// be passed, in turn, to this function so that their own subFolders can be detected and attached to them.
        /// E.g., first subFolder is called Folder1; it has 3 subfolders, Level21, Level22, and Level23. Level21 has two subfolders:
        /// Level31 and Level32.
        /// Folder1 will be passed to this function. Before execution is complete, Level21 will be passed to this function, then 
        /// Level31 and Level32 will be passed to it. Once the function has completed working on Level31 and Level32, it will 
        /// complete work on Level21. It will then complete work on Level22 and Level23 and finally complete on Folder1. By that time,
        /// all of Level21, Level22 and Level23 will be noted as subfolders of Folder1. Level21 (itself inside Folder1)
        /// will contain Level31 and Level32.
        /// </summary>
        /// <param name="inputFolder">This is subfolder object</param>
        /// <param name="path">This is the full file path corresponding to the subfolder object</param>
        
        public void populateSubFolders(SubFolder inputFolder, string path)
        {
            if (path.IndexOf("bu of electronic filing") < 0 &&
                path.IndexOf("Not for Circulation") < 0)
            {
                int numberOfSubFolders = 0;
                string[] subFolders;
                SubFolder newFolder;
                int slashPosition;
                string currentFolderName;
                string parentFolderName;
                int currentFolderDepth;
                try
                {
                    subFolders = Directory.GetDirectories(path, "*", SearchOption.TopDirectoryOnly);
                    numberOfSubFolders = subFolders.Length;
                    int i = 0;
                    if (numberOfSubFolders > 0)
                    {
                        for (i = 0; i < numberOfSubFolders; i++)
                        {
                            slashPosition = subFolders[i].LastIndexOf("\\");
                            currentFolderName = subFolders[i].Substring(slashPosition + 1);
                            if (currentFolderName.IndexOf(",") != -1)
                            {
                                currentFolderName.Replace(",", "','");
                            }
                            //for folder "T:\\BU - BUDGET AND FALL ECONOMIC STATEMENTS\\2021\\01 - Binder & Lock Up"
                            //depth is 3
                            currentFolderDepth = subFolders[i].Count(s => s == '\\');
                            currentFolderDepth = currentFolderDepth - Program.slashCorrection;
                            parentFolderName = subFolders[i].Substring(0, slashPosition);
                            slashPosition = parentFolderName.LastIndexOf("\\");
                            parentFolderName = parentFolderName.Substring(slashPosition + 1);
                            newFolder = inputFolder.createDeeperFolder(currentFolderName, currentFolderDepth, parentFolderName);
                            populateSubFolders(newFolder, subFolders[i]);
                        }
                    }
                    if (i == numberOfSubFolders)
                    {
                        Form1.currentSubFolderDone = true;
                    }
                }
                catch (System.Exception e)
                {
                    Form1.currentSubFolderDone = false;
                    /*
                    string message = "Error in collecting folder names. Clock OK and try again. If the problem persists," +
                        "please tell Mitch.";
                    message += "  System error message: " + e.ToString();
                    string caption = "Error collecting folder names";
                    MessageBoxButtons btns = MessageBoxButtons.OK;
                    DialogResult userResponse = MessageBox.Show(message, caption, btns);
                    if (userResponse == System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }
                    */

                    /*
                     * sample error :
                     * System.IO.DirectoryNotFoundException
  HResult=0x80070003
  Message=Could not find a part of the path 'T:\MD - MODELLING AND DATA\01 - Provdata and Models\
                    01 - General Series\01 - General\Benefit Calculator\Old versions 2012 Benefit Calculator\2011 Benefit Calculator base case for OCB disabled single parent\Benefit Calculator 2011\themes\pepper-grinder\images'.
                     * 
                     */
                }
            }
        }

    }
}
