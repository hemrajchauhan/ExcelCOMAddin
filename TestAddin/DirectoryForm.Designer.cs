namespace TestAddin
{
    partial class DirectoryForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DirectoryForm));
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.comboBoxExtension = new System.Windows.Forms.ComboBox();
            this.buttonExtension = new System.Windows.Forms.Button();
            this.buttonFileList = new System.Windows.Forms.Button();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.textBoxFileCount = new System.Windows.Forms.TextBox();
            this.buttonFileCount = new System.Windows.Forms.Button();
            this.buttonFolderList = new System.Windows.Forms.Button();
            this.button0KbFileList = new System.Windows.Forms.Button();
            this.buttonEmptyDirectoryList = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxPath
            // 
            this.textBoxPath.Location = new System.Drawing.Point(23, 25);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(395, 20);
            this.textBoxPath.TabIndex = 0;
            // 
            // comboBoxExtension
            // 
            this.comboBoxExtension.FormattingEnabled = true;
            this.comboBoxExtension.Location = new System.Drawing.Point(26, 76);
            this.comboBoxExtension.Name = "comboBoxExtension";
            this.comboBoxExtension.Size = new System.Drawing.Size(117, 21);
            this.comboBoxExtension.TabIndex = 1;
            // 
            // buttonExtension
            // 
            this.buttonExtension.Location = new System.Drawing.Point(193, 74);
            this.buttonExtension.Name = "buttonExtension";
            this.buttonExtension.Size = new System.Drawing.Size(98, 23);
            this.buttonExtension.TabIndex = 2;
            this.buttonExtension.Text = "Get Extensions";
            this.buttonExtension.UseVisualStyleBackColor = true;
            this.buttonExtension.Click += new System.EventHandler(this.buttonExtension_Click);
            // 
            // buttonFileList
            // 
            this.buttonFileList.Location = new System.Drawing.Point(341, 74);
            this.buttonFileList.Name = "buttonFileList";
            this.buttonFileList.Size = new System.Drawing.Size(77, 23);
            this.buttonFileList.TabIndex = 3;
            this.buttonFileList.Text = "File List";
            this.buttonFileList.UseVisualStyleBackColor = true;
            this.buttonFileList.Click += new System.EventHandler(this.buttonFileList_Click);
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(450, 25);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(75, 20);
            this.buttonBrowse.TabIndex = 4;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // textBoxFileCount
            // 
            this.textBoxFileCount.Location = new System.Drawing.Point(26, 128);
            this.textBoxFileCount.Name = "textBoxFileCount";
            this.textBoxFileCount.Size = new System.Drawing.Size(117, 20);
            this.textBoxFileCount.TabIndex = 5;
            // 
            // buttonFileCount
            // 
            this.buttonFileCount.Location = new System.Drawing.Point(193, 127);
            this.buttonFileCount.Name = "buttonFileCount";
            this.buttonFileCount.Size = new System.Drawing.Size(98, 21);
            this.buttonFileCount.TabIndex = 6;
            this.buttonFileCount.Text = "File Count";
            this.buttonFileCount.UseVisualStyleBackColor = true;
            this.buttonFileCount.Click += new System.EventHandler(this.buttonFileCount_Click);
            // 
            // buttonFolderList
            // 
            this.buttonFolderList.Location = new System.Drawing.Point(341, 128);
            this.buttonFolderList.Name = "buttonFolderList";
            this.buttonFolderList.Size = new System.Drawing.Size(77, 21);
            this.buttonFolderList.TabIndex = 7;
            this.buttonFolderList.Text = "Folder List";
            this.buttonFolderList.UseVisualStyleBackColor = true;
            this.buttonFolderList.Click += new System.EventHandler(this.buttonFolderList_Click);
            // 
            // button0KbFileList
            // 
            this.button0KbFileList.Location = new System.Drawing.Point(450, 74);
            this.button0KbFileList.Name = "button0KbFileList";
            this.button0KbFileList.Size = new System.Drawing.Size(84, 23);
            this.button0KbFileList.TabIndex = 8;
            this.button0KbFileList.Text = "0 Kb \\ Hidden";
            this.button0KbFileList.UseVisualStyleBackColor = true;
            this.button0KbFileList.Click += new System.EventHandler(this.button0KbFileList_Click);
            // 
            // buttonEmptyDirectoryList
            // 
            this.buttonEmptyDirectoryList.Location = new System.Drawing.Point(450, 125);
            this.buttonEmptyDirectoryList.Name = "buttonEmptyDirectoryList";
            this.buttonEmptyDirectoryList.Size = new System.Drawing.Size(84, 23);
            this.buttonEmptyDirectoryList.TabIndex = 9;
            this.buttonEmptyDirectoryList.Text = "Empty Dir List";
            this.buttonEmptyDirectoryList.UseVisualStyleBackColor = true;
            this.buttonEmptyDirectoryList.Click += new System.EventHandler(this.buttonEmptyDirectoryList_Click);
            // 
            // DirectoryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(555, 176);
            this.Controls.Add(this.buttonEmptyDirectoryList);
            this.Controls.Add(this.button0KbFileList);
            this.Controls.Add(this.buttonFolderList);
            this.Controls.Add(this.buttonFileCount);
            this.Controls.Add(this.textBoxFileCount);
            this.Controls.Add(this.buttonBrowse);
            this.Controls.Add(this.buttonFileList);
            this.Controls.Add(this.buttonExtension);
            this.Controls.Add(this.comboBoxExtension);
            this.Controls.Add(this.textBoxPath);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DirectoryForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Directory";
            this.Load += new System.EventHandler(this.DirectoryForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.ComboBox comboBoxExtension;
        private System.Windows.Forms.Button buttonExtension;
        private System.Windows.Forms.Button buttonFileList;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.TextBox textBoxFileCount;
        private System.Windows.Forms.Button buttonFileCount;
        private System.Windows.Forms.Button buttonFolderList;
        private System.Windows.Forms.Button button0KbFileList;
        private System.Windows.Forms.Button buttonEmptyDirectoryList;
    }
}