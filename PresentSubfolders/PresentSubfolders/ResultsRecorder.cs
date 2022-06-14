using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Threading;

namespace PresentSubfolders
{
    class ResultsRecorder
    {
        BackgroundWorker _worker;
        int _depthLevelsToDisplay;

        public ResultsRecorder(BackgroundWorker worker)
        {
            _worker = worker;
        }

        public int depthLevelsToDisplay
        {
            set { _depthLevelsToDisplay = value; }
        }


        /// <summary>
        /// This function does the work of writing the information about each individual subfolder. That information
        /// is kept in the string that will form the CSV.
        /// This is a recursive function, so if a subfolder that is passed to this function itself has next-level subfolders, 
        /// this function will call itself, passing in those next-level subfolders.
        /// </summary>
        /// <param name="sfIn">subfolder whose information is to be written to the csv</param>
        /// <param name="csvIn">The csv string (passed by reference)</param>
        private void recordSubFoldersInCSVString(SubFolder sfIn, ref StringBuilder csvIn)
        {
            string lineComposer;
            string fixedFolderName;
            for (int i = 0; i < sfIn.subFolders.Count; i++)
            {//record the items at the top level of the queue of subfolders that was passed in
                if (sfIn.subFolders.ElementAt(i).sfLevel <= _depthLevelsToDisplay)
                {//only record the folder if it's within the levels to be displayed
                    //add commmas to reflect depth level of the subfolder; these will turn into columns when CSV is converted to excel.
                    lineComposer = new String(',', sfIn.subFolders.ElementAt(i).sfLevel - 1);
                    fixedFolderName = sfIn.subFolders.ElementAt(i).sfName;
                    //if the subfolder's name is a year, add approstrophe, so that excel handles it as text
                    fixedFolderName = System.Text.RegularExpressions.Regex.Replace(fixedFolderName,
                        @"^\s*\d{4}\s*$", @"'$&'");
                    lineComposer += fixedFolderName;
                    csvIn.AppendLine(lineComposer);
                    if (sfIn.subFolders.ElementAt(i).subFolders.Count > 0)
                    {//if there are more subFolders within this one, record all of them
                        recordSubFoldersInCSVString(sfIn.subFolders.ElementAt(i), ref csvIn);
                    }
                }
            }
        }

        //!! the following function and its helpers need to be pushed off to 
        //a seaparate class, with a backgroundworker as go between from that class to this UI.
        //bgworker has been added for that
        //!!you'll need a way of triggering the appropriate all done panel. maybe just have the same one. or you could have different backgroundworkers
        //for word and excel and the completed method of each can set the appropriate panel. But you can probably have some UI-level variable to say
        //whether you're working with word or excel and use that in a generic completed method of the bgworker to detect which panel to display.

        /// <summary>
        /// This is the main function to control output to a word or excel file.
        /// Each subfolder that exists gets passed to recordSubFolders(), which creates a record, in the string that 
        /// will form the CSV, for that individual subFolder.
        /// 
        /// /// </summary>
        public void recordResultsToExcelOrWord(string outputMode)
        {            //output results to a file
            Excel.Application oXL;
            Workbook wb;
            Worksheet ws;
            Excel.Range oRng;
            string excelFileName = "";
            oXL = new Excel.Application();
            wb = oXL.Workbooks.Add(System.Reflection.Missing.Value);
            if (outputMode == "word")
            {//if we're exporting to word, set the depth level for display at 3
                _depthLevelsToDisplay = 3;
            }
            try
            {
                var csv = new StringBuilder();
                string row;
                string fixedFolderName;
                string baseFileName = "";
                //populate this whole thing
                for (int i = 0; i < Program.centralFiles.subFolders.Count; i++)
                {
                    //this is just the top level of subfolders
                    //record the top level subfolder
                    fixedFolderName = Program.centralFiles.subFolders.ElementAt(i).sfName;
                    //if the subfolder's name is a year, add approstrophe, so that excel handles it as text
                    fixedFolderName = System.Text.RegularExpressions.Regex.Replace(fixedFolderName,
                        @"^\s*\d{4}\s*$", @"'$&'");
                    row = fixedFolderName;
                    csv.AppendLine(row);
                    if (Program.centralFiles.subFolders.ElementAt(i).subFolders.Count > 0)
                    {//this subfolder has subfolders of its own, so go record them
                        recordSubFoldersInCSVString(Program.centralFiles.subFolders.ElementAt(i), ref csv);
                    }
                }

                Thread t = new Thread((ThreadStart)(() =>
                {
                    System.Windows.Forms.SaveFileDialog sa = new System.Windows.Forms.SaveFileDialog();
                    sa.InitialDirectory = @"C:\";
                    sa.Title = "Where to save the folder list";
                    if (outputMode == "word")
                    {
                        sa.DefaultExt = "docx";
                        sa.Filter = "Word Files(.docx)|*.docx";
                    }
                    else
                    {
                        sa.DefaultExt = "xlsx";
                        sa.Filter = "Excel Files(.xlsx)|*.xlsx";
                    }
                    sa.RestoreDirectory = true;
                    if (sa.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        baseFileName = sa.FileName;
                    }
                }));
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                //save as CSV, then open that in excel.
                string csvFileName = baseFileName + ".csv";
                excelFileName = baseFileName.Replace(".docx", ".xlsx"); //if it's docx file, the excel filename will have a .xlsx; if it's an excel file,
                                                                           //it will already have a .xlsx on it, and this line will do nothing
                string wordFileName;

                File.WriteAllText(csvFileName, csv.ToString());
                //open in excel, format, and save as excel file
                wb = oXL.Workbooks.Open(csvFileName);
                ws = wb.Worksheets[1];
                ws.Columns.ColumnWidth = 2;
                ws.Range["A1:B1"].EntireColumn.Font.Size = 12;
                ws.Range["A1:B1"].EntireColumn.ColumnWidth = 1.5;
                oRng = ws.Columns[1];
                oRng.EntireColumn.Font.Bold = true;
                oRng = ws.Columns[2];
                oRng.EntireColumn.Font.Underline = true;

                wb.SaveAs(excelFileName, XlFileFormat.xlOpenXMLWorkbook,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                if (outputMode == "excel")
                { //if we're outputting excel, save the file and display the results in excel panel
                      //!!the following UI stuff should be moved back to the form
                      //you'll need a way to report the text for the link back to that UI
                      //or have the popup for the file name in the UI function, send the name to this helper
                    Form1.excelFileLinkText = excelFileName;
                }
                //get rid of the csv file
                File.Delete(csvFileName);

                if (outputMode == "word")
                {

                    wordFileName = baseFileName;
                    object oMissing = System.Reflection.Missing.Value;
                    Word._Application wrd = new Word.Application();
                    wrd.Visible = false;
                    Word._Document doc;
                    doc = wrd.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    double report = 0;//this and max (defined just below) will be used in reporting progress to UI

                    try
                    {
                        ws.Range["A1:C1"].EntireColumn.Copy();
                        doc.ActiveWindow.Selection.PasteExcelTable(false, false, false);
                        //next
                        t = new Thread((ThreadStart)(() =>
                        {
                            Clipboard.Clear();
                        }));
                        t.SetApartmentState(ApartmentState.STA);
                        t.Start();
                        t.Join();
                        Word.Table table = doc.Tables[1];
                        double max = table.Rows.Count + 2;//this will be used for reporting progress of work back to UI
                        table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                        //the following loop sets spans and widths of columns
                        for (int rowInTable = 1; rowInTable <= table.Rows.Count; rowInTable++)
                        {
                            if (table.Rows[rowInTable].Cells.Count == 3 && table.Cell(rowInTable, 2).Range.Text == "\r\a" &&
                                table.Cell(rowInTable, 2).Width >23)
                            {
                                table.Cell(rowInTable, 2).Width = 23;
                            }
                            report = System.Convert.ToDouble(rowInTable) / max;//calculate progress of work
                            report = report * 100;
                            _worker.ReportProgress(System.Convert.ToInt32(report));//report progress to UI
                        }
                        //clean up formatting
                        Word.Range location = doc.Range(doc.Content.Start, doc.Content.End);
                        //                    location.Font.Size = 11;
                        location.Paragraphs.SpaceAfter = 0;
                        doc.PageSetup.TextColumns.SetCount(2);
                        doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                        doc.SaveAs2(wordFileName);
                        //doc.Close();
                        Form1.wordFileLinkText = wordFileName;
                    }
                    catch (Exception e)
                    {
                        //error in the attempt to write excel file
                        string message = "Error writing to Word. Clock OK and try again. If the problem persists," +
                            "please tell Mitch.";
                        message += "  System error message: " + e.ToString();
                        string caption = "Word writing error";
                        MessageBoxButtons btns = MessageBoxButtons.OK;
                        DialogResult userResponse = MessageBox.Show(message, caption, btns);
                        if (userResponse == System.Windows.Forms.DialogResult.OK)
                        {
                            //!!!maybe you need to throw an error here to report to any calling function
                            return;
                        }
                    }
                    finally
                    {
                        t = new Thread((ThreadStart)(() =>
                        {
                            Clipboard.Clear();
                        }));
                        t.SetApartmentState(ApartmentState.STA);
                        t.Start();
                        t.Join();
                        //doc.Close();
                        wrd.Quit();
                    }
                }
                
            }
            catch (Exception e)
            {
                //error in the attempt to write excel file
                string message = "Error writing to excel. Clock OK and try again. If the problem persists," +
                    "please tell Mitch.";
                message += "  System error message: " + e.ToString();
                string caption = "Excel writing error";
                MessageBoxButtons btns = MessageBoxButtons.OK;
                DialogResult userResponse = MessageBox.Show(message, caption, btns);
                if (userResponse == System.Windows.Forms.DialogResult.OK)
                {
                    //!!!maybe you need to throw an error here to report to any calling function
                    return;
                }
            }
            finally
            {
                wb.Close();
                oXL.Quit();
                if (outputMode == "word")
                {
                    File.Delete(excelFileName);
                }
            }
        }



    }
}
