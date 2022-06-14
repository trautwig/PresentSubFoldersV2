
namespace PresentSubfolders
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.doneLabel = new System.Windows.Forms.Label();
            this.excelFileLink = new System.Windows.Forms.LinkLabel();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.createExcel = new System.Windows.Forms.Button();
            this.selectLevelText = new System.Windows.Forms.Label();
            this.depthLevelText = new System.Windows.Forms.Label();
            this.depthLevelsToDisplay = new System.Windows.Forms.NumericUpDown();
            this.panel1 = new System.Windows.Forms.Panel();
            this.levelDisplayGroup = new System.Windows.Forms.Panel();
            this.createWordButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.badResultsPanel = new System.Windows.Forms.Panel();
            this.badResultsLabel = new System.Windows.Forms.Label();
            this.excelResultsDonePanel = new System.Windows.Forms.Panel();
            this.workingLabel = new System.Windows.Forms.Label();
            this.workingOnFilesPanel = new System.Windows.Forms.Panel();
            this.folderRecordingProgressBar = new System.Windows.Forms.ProgressBar();
            this.workingOnWordPanel = new System.Windows.Forms.Panel();
            this.workingOnWordProgressBar = new System.Windows.Forms.ProgressBar();
            this.workingOnWordLabel = new System.Windows.Forms.Label();
            this.wordResultsDonePanel = new System.Windows.Forms.Panel();
            this.wordFileLink = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.bgWorkerFolderSearch = new System.ComponentModel.BackgroundWorker();
            this.bgWorkerResultsToExcelOrWord = new System.ComponentModel.BackgroundWorker();
            ((System.ComponentModel.ISupportInitialize)(this.depthLevelsToDisplay)).BeginInit();
            this.panel1.SuspendLayout();
            this.levelDisplayGroup.SuspendLayout();
            this.badResultsPanel.SuspendLayout();
            this.excelResultsDonePanel.SuspendLayout();
            this.workingOnFilesPanel.SuspendLayout();
            this.workingOnWordPanel.SuspendLayout();
            this.wordResultsDonePanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // doneLabel
            // 
            this.doneLabel.AutoSize = true;
            this.doneLabel.Location = new System.Drawing.Point(3, 10);
            this.doneLabel.Name = "doneLabel";
            this.doneLabel.Size = new System.Drawing.Size(169, 13);
            this.doneLabel.TabIndex = 1;
            this.doneLabel.Text = "All done! Go check your Excel file:";
            // 
            // excelFileLink
            // 
            this.excelFileLink.AutoSize = true;
            this.excelFileLink.Location = new System.Drawing.Point(178, 10);
            this.excelFileLink.Name = "excelFileLink";
            this.excelFileLink.Size = new System.Drawing.Size(0, 13);
            this.excelFileLink.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(6, 144);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Where are the folders to list?";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.SelectFolders);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(570, 68);
            this.label2.TabIndex = 3;
            this.label2.Text = resources.GetString("label2.Text");
            // 
            // createExcel
            // 
            this.createExcel.Location = new System.Drawing.Point(9, 45);
            this.createExcel.Name = "createExcel";
            this.createExcel.Size = new System.Drawing.Size(89, 23);
            this.createExcel.TabIndex = 7;
            this.createExcel.Text = "Create Excel File";
            this.createExcel.UseVisualStyleBackColor = true;
            this.createExcel.Click += new System.EventHandler(this.createExcel_Click);
            // 
            // selectLevelText
            // 
            this.selectLevelText.AutoSize = true;
            this.selectLevelText.Location = new System.Drawing.Point(3, 23);
            this.selectLevelText.Name = "selectLevelText";
            this.selectLevelText.Size = new System.Drawing.Size(185, 13);
            this.selectLevelText.TabIndex = 6;
            this.selectLevelText.Text = "Select the number of levels to display:";
            // 
            // depthLevelText
            // 
            this.depthLevelText.Location = new System.Drawing.Point(3, 0);
            this.depthLevelText.Name = "depthLevelText";
            this.depthLevelText.Size = new System.Drawing.Size(306, 18);
            this.depthLevelText.TabIndex = 4;
            this.depthLevelText.Text = "The folders found went to a depth of";
            // 
            // depthLevelsToDisplay
            // 
            this.depthLevelsToDisplay.Location = new System.Drawing.Point(194, 23);
            this.depthLevelsToDisplay.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.depthLevelsToDisplay.Name = "depthLevelsToDisplay";
            this.depthLevelsToDisplay.Size = new System.Drawing.Size(49, 20);
            this.depthLevelsToDisplay.TabIndex = 5;
            this.depthLevelsToDisplay.Value = new decimal(new int[] {
            4,
            0,
            0,
            0});
            this.depthLevelsToDisplay.ValueChanged += new System.EventHandler(this.depthLevelsToDisplay_ValueChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.levelDisplayGroup);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.badResultsPanel);
            this.panel1.Location = new System.Drawing.Point(19, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(577, 212);
            this.panel1.TabIndex = 0;
            // 
            // levelDisplayGroup
            // 
            this.levelDisplayGroup.Controls.Add(this.createWordButton);
            this.levelDisplayGroup.Controls.Add(this.label1);
            this.levelDisplayGroup.Controls.Add(this.createExcel);
            this.levelDisplayGroup.Controls.Add(this.depthLevelText);
            this.levelDisplayGroup.Controls.Add(this.selectLevelText);
            this.levelDisplayGroup.Controls.Add(this.depthLevelsToDisplay);
            this.levelDisplayGroup.Location = new System.Drawing.Point(162, 127);
            this.levelDisplayGroup.Name = "levelDisplayGroup";
            this.levelDisplayGroup.Size = new System.Drawing.Size(391, 82);
            this.levelDisplayGroup.TabIndex = 5;
            // 
            // createWordButton
            // 
            this.createWordButton.Location = new System.Drawing.Point(142, 45);
            this.createWordButton.Name = "createWordButton";
            this.createWordButton.Size = new System.Drawing.Size(145, 23);
            this.createWordButton.TabIndex = 9;
            this.createWordButton.Text = "Create 3 Level Word File";
            this.createWordButton.UseVisualStyleBackColor = true;
            this.createWordButton.Click += new System.EventHandler(this.createWord_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(109, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "OR";
            // 
            // badResultsPanel
            // 
            this.badResultsPanel.Controls.Add(this.badResultsLabel);
            this.badResultsPanel.Location = new System.Drawing.Point(166, 125);
            this.badResultsPanel.Name = "badResultsPanel";
            this.badResultsPanel.Size = new System.Drawing.Size(358, 83);
            this.badResultsPanel.TabIndex = 6;
            // 
            // badResultsLabel
            // 
            this.badResultsLabel.Location = new System.Drawing.Point(16, 16);
            this.badResultsLabel.Name = "badResultsLabel";
            this.badResultsLabel.Size = new System.Drawing.Size(325, 48);
            this.badResultsLabel.TabIndex = 0;
            this.badResultsLabel.Text = "Failed to capture all folders. This may be the result of a network interruption. " +
    "Consdider trying again.";
            // 
            // excelResultsDonePanel
            // 
            this.excelResultsDonePanel.Controls.Add(this.doneLabel);
            this.excelResultsDonePanel.Controls.Add(this.excelFileLink);
            this.excelResultsDonePanel.Location = new System.Drawing.Point(19, 266);
            this.excelResultsDonePanel.Name = "excelResultsDonePanel";
            this.excelResultsDonePanel.Size = new System.Drawing.Size(630, 37);
            this.excelResultsDonePanel.TabIndex = 4;
            // 
            // workingLabel
            // 
            this.workingLabel.AutoSize = true;
            this.workingLabel.Location = new System.Drawing.Point(1, 9);
            this.workingLabel.Name = "workingLabel";
            this.workingLabel.Size = new System.Drawing.Size(364, 13);
            this.workingLabel.TabIndex = 5;
            this.workingLabel.Text = "Working ... this may take a while for a large set of folders on a network drive";
            // 
            // workingOnFilesPanel
            // 
            this.workingOnFilesPanel.Controls.Add(this.folderRecordingProgressBar);
            this.workingOnFilesPanel.Controls.Add(this.workingLabel);
            this.workingOnFilesPanel.Location = new System.Drawing.Point(19, 230);
            this.workingOnFilesPanel.Name = "workingOnFilesPanel";
            this.workingOnFilesPanel.Size = new System.Drawing.Size(552, 36);
            this.workingOnFilesPanel.TabIndex = 6;
            // 
            // folderRecordingProgressBar
            // 
            this.folderRecordingProgressBar.Location = new System.Drawing.Point(370, 6);
            this.folderRecordingProgressBar.Name = "folderRecordingProgressBar";
            this.folderRecordingProgressBar.Size = new System.Drawing.Size(100, 23);
            this.folderRecordingProgressBar.TabIndex = 6;
            // 
            // workingOnWordPanel
            // 
            this.workingOnWordPanel.Controls.Add(this.workingOnWordProgressBar);
            this.workingOnWordPanel.Controls.Add(this.workingOnWordLabel);
            this.workingOnWordPanel.Location = new System.Drawing.Point(19, 229);
            this.workingOnWordPanel.Name = "workingOnWordPanel";
            this.workingOnWordPanel.Size = new System.Drawing.Size(341, 37);
            this.workingOnWordPanel.TabIndex = 7;
            // 
            // workingOnWordProgressBar
            // 
            this.workingOnWordProgressBar.Location = new System.Drawing.Point(208, 9);
            this.workingOnWordProgressBar.Name = "workingOnWordProgressBar";
            this.workingOnWordProgressBar.Size = new System.Drawing.Size(100, 23);
            this.workingOnWordProgressBar.TabIndex = 1;
            // 
            // workingOnWordLabel
            // 
            this.workingOnWordLabel.AutoSize = true;
            this.workingOnWordLabel.Location = new System.Drawing.Point(3, 14);
            this.workingOnWordLabel.Name = "workingOnWordLabel";
            this.workingOnWordLabel.Size = new System.Drawing.Size(189, 13);
            this.workingOnWordLabel.TabIndex = 0;
            this.workingOnWordLabel.Text = "Working ... this can be slow with Word";
            // 
            // wordResultsDonePanel
            // 
            this.wordResultsDonePanel.Controls.Add(this.wordFileLink);
            this.wordResultsDonePanel.Controls.Add(this.label3);
            this.wordResultsDonePanel.Location = new System.Drawing.Point(19, 268);
            this.wordResultsDonePanel.Name = "wordResultsDonePanel";
            this.wordResultsDonePanel.Size = new System.Drawing.Size(630, 33);
            this.wordResultsDonePanel.TabIndex = 8;
            // 
            // wordFileLink
            // 
            this.wordFileLink.AutoSize = true;
            this.wordFileLink.Location = new System.Drawing.Point(170, 10);
            this.wordFileLink.Name = "wordFileLink";
            this.wordFileLink.Size = new System.Drawing.Size(55, 13);
            this.wordFileLink.TabIndex = 1;
            this.wordFileLink.TabStop = true;
            this.wordFileLink.Text = "linkLabel1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(169, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "All done! Go check your Word file:";
            // 
            // panel2
            // 
            this.panel2.Location = new System.Drawing.Point(611, 20);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(48, 210);
            this.panel2.TabIndex = 9;
            // 
            // bgWorkerFolderSearch
            // 
            this.bgWorkerFolderSearch.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgWorkerFolderSearch_DoWork);
            this.bgWorkerFolderSearch.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgWorkerFolderSearch_ProgressChanged);
            this.bgWorkerFolderSearch.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgWorkerFolderSearch_RunWorkerCompleted);
            // 
            // bgWorkerResultsToExcelOrWord
            // 
            this.bgWorkerResultsToExcelOrWord.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgWorkerResultsToExcelOrWord_DoWork);
            this.bgWorkerResultsToExcelOrWord.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgWorkerResultsToExcelOrWord_ProgressChanged);
            this.bgWorkerResultsToExcelOrWord.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgWorkerResultsToExcelOrWord_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(663, 345);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.wordResultsDonePanel);
            this.Controls.Add(this.workingOnWordPanel);
            this.Controls.Add(this.workingOnFilesPanel);
            this.Controls.Add(this.excelResultsDonePanel);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.Text = "Records Recording";
            ((System.ComponentModel.ISupportInitialize)(this.depthLevelsToDisplay)).EndInit();
            this.panel1.ResumeLayout(false);
            this.levelDisplayGroup.ResumeLayout(false);
            this.levelDisplayGroup.PerformLayout();
            this.badResultsPanel.ResumeLayout(false);
            this.excelResultsDonePanel.ResumeLayout(false);
            this.excelResultsDonePanel.PerformLayout();
            this.workingOnFilesPanel.ResumeLayout(false);
            this.workingOnFilesPanel.PerformLayout();
            this.workingOnWordPanel.ResumeLayout(false);
            this.workingOnWordPanel.PerformLayout();
            this.wordResultsDonePanel.ResumeLayout(false);
            this.wordResultsDonePanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label doneLabel;
        private System.Windows.Forms.LinkLabel excelFileLink;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button createExcel;
        private System.Windows.Forms.Label selectLevelText;
        private System.Windows.Forms.Label depthLevelText;
        private System.Windows.Forms.NumericUpDown depthLevelsToDisplay;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel excelResultsDonePanel;
        private System.Windows.Forms.Panel levelDisplayGroup;
        private System.Windows.Forms.Label workingLabel;
        private System.Windows.Forms.Panel badResultsPanel;
        private System.Windows.Forms.Label badResultsLabel;
        private System.Windows.Forms.Panel workingOnFilesPanel;
        private System.Windows.Forms.ProgressBar folderRecordingProgressBar;
        private System.Windows.Forms.Button createWordButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel workingOnWordPanel;
        private System.Windows.Forms.ProgressBar workingOnWordProgressBar;
        private System.Windows.Forms.Label workingOnWordLabel;
        private System.Windows.Forms.Panel wordResultsDonePanel;
        private System.Windows.Forms.LinkLabel wordFileLink;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel panel2;
        private System.ComponentModel.BackgroundWorker bgWorkerFolderSearch;
        private System.ComponentModel.BackgroundWorker bgWorkerResultsToExcelOrWord;
    }
}

