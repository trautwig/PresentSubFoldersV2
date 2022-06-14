using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PresentSubfolders
{
    static class Program
    {
        public static int maxLevel;//will track the maximum subFolder depth level.
        public static string searchString;
        public static int slashCorrection;
        public static TopLevelFolder centralFiles;
        //centralFiles is the subFolder that will contain everything


        // this allows the subfolders to record their depth by looking at their relation to whatever folder is
        //used as base folder for collecting subfolder information

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
