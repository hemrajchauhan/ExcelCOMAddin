namespace TestAddin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.buttonDeleteEmptyDirectory = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.buttonCreateFolder = this.Factory.CreateRibbonButton();
            this.buttonDeleteDirectory = this.Factory.CreateRibbonButton();
            this.buttonMoveDirectory = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.buttonCopyFiles = this.Factory.CreateRibbonButton();
            this.buttonMoveFiles = this.Factory.CreateRibbonButton();
            this.buttonReplaceFiles = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.buttonRenameFiles = this.Factory.CreateRibbonButton();
            this.buttonUnhideFiles = this.Factory.CreateRibbonButton();
            this.buttonDeleteFiles = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonUpper = this.Factory.CreateRibbonButton();
            this.buttonLower = this.Factory.CreateRibbonButton();
            this.buttonTitle = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.buttonTrimSpace = this.Factory.CreateRibbonButton();
            this.buttonTrimEnter = this.Factory.CreateRibbonButton();
            this.buttonJoin = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.buttonLeftNth = this.Factory.CreateRibbonButton();
            this.buttonMidNth = this.Factory.CreateRibbonButton();
            this.buttonRightNth = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.buttonHyperlink = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.buttonCombineExcel = this.Factory.CreateRibbonButton();
            this.Sheets = this.Factory.CreateRibbonGroup();
            this.buttonCombineSheets = this.Factory.CreateRibbonButton();
            this.buttonSplitSheets = this.Factory.CreateRibbonButton();
            this.Dropbox = this.Factory.CreateRibbonGroup();
            this.buttonFileList = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.Sheets.SuspendLayout();
            this.Dropbox.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.Sheets);
            this.tab1.Groups.Add(this.Dropbox);
            this.tab1.Label = "Others";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.AfterOfficeId("View");
            // 
            // group1
            // 
            ribbonDialogLauncherImpl1.Enabled = false;
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.buttonDeleteEmptyDirectory);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.buttonCreateFolder);
            this.group1.Items.Add(this.buttonDeleteDirectory);
            this.group1.Items.Add(this.buttonMoveDirectory);
            this.group1.Label = "Directory";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "File Browser";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // buttonDeleteEmptyDirectory
            // 
            this.buttonDeleteEmptyDirectory.Image = ((System.Drawing.Image)(resources.GetObject("buttonDeleteEmptyDirectory.Image")));
            this.buttonDeleteEmptyDirectory.Label = "Delete Empty";
            this.buttonDeleteEmptyDirectory.Name = "buttonDeleteEmptyDirectory";
            this.buttonDeleteEmptyDirectory.ScreenTip = "Delete empty folders";
            this.buttonDeleteEmptyDirectory.ShowImage = true;
            this.buttonDeleteEmptyDirectory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDeleteEmptyDirectory_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // buttonCreateFolder
            // 
            this.buttonCreateFolder.Image = ((System.Drawing.Image)(resources.GetObject("buttonCreateFolder.Image")));
            this.buttonCreateFolder.Label = "Create";
            this.buttonCreateFolder.Name = "buttonCreateFolder";
            this.buttonCreateFolder.ScreenTip = "Create folders";
            this.buttonCreateFolder.ShowImage = true;
            this.buttonCreateFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateFolder_Click);
            // 
            // buttonDeleteDirectory
            // 
            this.buttonDeleteDirectory.Image = ((System.Drawing.Image)(resources.GetObject("buttonDeleteDirectory.Image")));
            this.buttonDeleteDirectory.Label = "Delete";
            this.buttonDeleteDirectory.Name = "buttonDeleteDirectory";
            this.buttonDeleteDirectory.ScreenTip = "Delete folders";
            this.buttonDeleteDirectory.ShowImage = true;
            this.buttonDeleteDirectory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDeleteDirectory_Click);
            // 
            // buttonMoveDirectory
            // 
            this.buttonMoveDirectory.Image = ((System.Drawing.Image)(resources.GetObject("buttonMoveDirectory.Image")));
            this.buttonMoveDirectory.Label = "Move";
            this.buttonMoveDirectory.Name = "buttonMoveDirectory";
            this.buttonMoveDirectory.ShowImage = true;
            this.buttonMoveDirectory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMoveDirectory_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.buttonCopyFiles);
            this.group3.Items.Add(this.buttonMoveFiles);
            this.group3.Items.Add(this.buttonReplaceFiles);
            this.group3.Items.Add(this.separator2);
            this.group3.Items.Add(this.buttonRenameFiles);
            this.group3.Items.Add(this.buttonUnhideFiles);
            this.group3.Items.Add(this.buttonDeleteFiles);
            this.group3.Label = "Files";
            this.group3.Name = "group3";
            // 
            // buttonCopyFiles
            // 
            this.buttonCopyFiles.Image = ((System.Drawing.Image)(resources.GetObject("buttonCopyFiles.Image")));
            this.buttonCopyFiles.Label = "Copy";
            this.buttonCopyFiles.Name = "buttonCopyFiles";
            this.buttonCopyFiles.ScreenTip = "Copy files";
            this.buttonCopyFiles.ShowImage = true;
            this.buttonCopyFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCopyFiles_Click);
            // 
            // buttonMoveFiles
            // 
            this.buttonMoveFiles.Image = ((System.Drawing.Image)(resources.GetObject("buttonMoveFiles.Image")));
            this.buttonMoveFiles.Label = "Move";
            this.buttonMoveFiles.Name = "buttonMoveFiles";
            this.buttonMoveFiles.ScreenTip = "Move files";
            this.buttonMoveFiles.ShowImage = true;
            this.buttonMoveFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMoveFiles_Click);
            // 
            // buttonReplaceFiles
            // 
            this.buttonReplaceFiles.Image = ((System.Drawing.Image)(resources.GetObject("buttonReplaceFiles.Image")));
            this.buttonReplaceFiles.Label = "Replace";
            this.buttonReplaceFiles.Name = "buttonReplaceFiles";
            this.buttonReplaceFiles.ScreenTip = "Replace files";
            this.buttonReplaceFiles.ShowImage = true;
            this.buttonReplaceFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReplaceFiles_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // buttonRenameFiles
            // 
            this.buttonRenameFiles.Image = ((System.Drawing.Image)(resources.GetObject("buttonRenameFiles.Image")));
            this.buttonRenameFiles.Label = "Rename";
            this.buttonRenameFiles.Name = "buttonRenameFiles";
            this.buttonRenameFiles.ShowImage = true;
            this.buttonRenameFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRenameFiles_Click);
            // 
            // buttonUnhideFiles
            // 
            this.buttonUnhideFiles.Image = ((System.Drawing.Image)(resources.GetObject("buttonUnhideFiles.Image")));
            this.buttonUnhideFiles.Label = "Unhide";
            this.buttonUnhideFiles.Name = "buttonUnhideFiles";
            this.buttonUnhideFiles.ScreenTip = "Unhide files";
            this.buttonUnhideFiles.ShowImage = true;
            this.buttonUnhideFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUnhideFiles_Click);
            // 
            // buttonDeleteFiles
            // 
            this.buttonDeleteFiles.Image = ((System.Drawing.Image)(resources.GetObject("buttonDeleteFiles.Image")));
            this.buttonDeleteFiles.Label = "Delete";
            this.buttonDeleteFiles.Name = "buttonDeleteFiles";
            this.buttonDeleteFiles.ScreenTip = "Delete files";
            this.buttonDeleteFiles.ShowImage = true;
            this.buttonDeleteFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDeleteFiles_Click);
            // 
            // group2
            // 
            ribbonDialogLauncherImpl2.Enabled = false;
            this.group2.DialogLauncher = ribbonDialogLauncherImpl2;
            this.group2.Items.Add(this.buttonUpper);
            this.group2.Items.Add(this.buttonLower);
            this.group2.Items.Add(this.buttonTitle);
            this.group2.Items.Add(this.separator3);
            this.group2.Items.Add(this.buttonTrimSpace);
            this.group2.Items.Add(this.buttonTrimEnter);
            this.group2.Items.Add(this.buttonJoin);
            this.group2.Items.Add(this.separator4);
            this.group2.Items.Add(this.buttonLeftNth);
            this.group2.Items.Add(this.buttonMidNth);
            this.group2.Items.Add(this.buttonRightNth);
            this.group2.Items.Add(this.separator5);
            this.group2.Items.Add(this.buttonHyperlink);
            this.group2.Label = "Text";
            this.group2.Name = "group2";
            // 
            // buttonUpper
            // 
            this.buttonUpper.Image = ((System.Drawing.Image)(resources.GetObject("buttonUpper.Image")));
            this.buttonUpper.Label = "Upper";
            this.buttonUpper.Name = "buttonUpper";
            this.buttonUpper.ShowImage = true;
            this.buttonUpper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUpper_Click);
            // 
            // buttonLower
            // 
            this.buttonLower.Image = ((System.Drawing.Image)(resources.GetObject("buttonLower.Image")));
            this.buttonLower.Label = "Lower";
            this.buttonLower.Name = "buttonLower";
            this.buttonLower.ShowImage = true;
            this.buttonLower.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLower_Click);
            // 
            // buttonTitle
            // 
            this.buttonTitle.Image = ((System.Drawing.Image)(resources.GetObject("buttonTitle.Image")));
            this.buttonTitle.Label = "Title";
            this.buttonTitle.Name = "buttonTitle";
            this.buttonTitle.ShowImage = true;
            this.buttonTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTitle_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // buttonTrimSpace
            // 
            this.buttonTrimSpace.Label = "Trim";
            this.buttonTrimSpace.Name = "buttonTrimSpace";
            this.buttonTrimSpace.ScreenTip = "Trim spaces";
            this.buttonTrimSpace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTrimSpace_Click);
            // 
            // buttonTrimEnter
            // 
            this.buttonTrimEnter.Label = "Trim <-\'";
            this.buttonTrimEnter.Name = "buttonTrimEnter";
            this.buttonTrimEnter.ScreenTip = "Trim Enter";
            this.buttonTrimEnter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTrimEnter_Click);
            // 
            // buttonJoin
            // 
            this.buttonJoin.Label = "Join";
            this.buttonJoin.Name = "buttonJoin";
            this.buttonJoin.ScreenTip = "Join Strings";
            this.buttonJoin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonJoin_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // buttonLeftNth
            // 
            this.buttonLeftNth.Label = "Left Nth";
            this.buttonLeftNth.Name = "buttonLeftNth";
            this.buttonLeftNth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLeftNth_Click);
            // 
            // buttonMidNth
            // 
            this.buttonMidNth.Label = "Mid Nth";
            this.buttonMidNth.Name = "buttonMidNth";
            this.buttonMidNth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMidNth_Click);
            // 
            // buttonRightNth
            // 
            this.buttonRightNth.Label = "Right Nth";
            this.buttonRightNth.Name = "buttonRightNth";
            this.buttonRightNth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRightNth_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // buttonHyperlink
            // 
            this.buttonHyperlink.Image = ((System.Drawing.Image)(resources.GetObject("buttonHyperlink.Image")));
            this.buttonHyperlink.Label = "Hyperlink";
            this.buttonHyperlink.Name = "buttonHyperlink";
            this.buttonHyperlink.ShowImage = true;
            this.buttonHyperlink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonHyperlink_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.buttonCombineExcel);
            this.group4.Label = "Excel";
            this.group4.Name = "group4";
            // 
            // buttonCombineExcel
            // 
            this.buttonCombineExcel.Image = ((System.Drawing.Image)(resources.GetObject("buttonCombineExcel.Image")));
            this.buttonCombineExcel.Label = "Import";
            this.buttonCombineExcel.Name = "buttonCombineExcel";
            this.buttonCombineExcel.ShowImage = true;
            this.buttonCombineExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCombineExcel_Click);
            // 
            // Sheets
            // 
            this.Sheets.Items.Add(this.buttonCombineSheets);
            this.Sheets.Items.Add(this.buttonSplitSheets);
            this.Sheets.Label = "WorkSheet";
            this.Sheets.Name = "Sheets";
            // 
            // buttonCombineSheets
            // 
            this.buttonCombineSheets.Image = ((System.Drawing.Image)(resources.GetObject("buttonCombineSheets.Image")));
            this.buttonCombineSheets.Label = "Combine";
            this.buttonCombineSheets.Name = "buttonCombineSheets";
            this.buttonCombineSheets.ShowImage = true;
            this.buttonCombineSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCombineSheets_Click);
            // 
            // buttonSplitSheets
            // 
            this.buttonSplitSheets.Image = ((System.Drawing.Image)(resources.GetObject("buttonSplitSheets.Image")));
            this.buttonSplitSheets.Label = "Split";
            this.buttonSplitSheets.Name = "buttonSplitSheets";
            this.buttonSplitSheets.ShowImage = true;
            this.buttonSplitSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSplitSheets_Click);
            // 
            // Dropbox
            // 
            this.Dropbox.Items.Add(this.buttonFileList);
            this.Dropbox.Label = "Dropbox";
            this.Dropbox.Name = "Dropbox";
            // 
            // buttonFileList
            // 
            this.buttonFileList.Label = "File List";
            this.buttonFileList.Name = "buttonFileList";
            this.buttonFileList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFileList_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.Sheets.ResumeLayout(false);
            this.Sheets.PerformLayout();
            this.Dropbox.ResumeLayout(false);
            this.Dropbox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLower;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCopyFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMoveFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteDirectory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReplaceFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteEmptyDirectory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTrimSpace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUnhideFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonJoin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTrimEnter;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMoveDirectory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCombineSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSplitSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLeftNth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMidNth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRightNth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonHyperlink;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Sheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRenameFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        //internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCombineExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Dropbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFileList;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
