namespace TestAddin
{
    partial class CompareForm
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
            this.comboBoxSheet1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBoxSheet2 = new System.Windows.Forms.ComboBox();
            this.listBoxSheet1 = new System.Windows.Forms.ListBox();
            this.listBoxSheet2 = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // comboBoxSheet1
            // 
            this.comboBoxSheet1.FormattingEnabled = true;
            this.comboBoxSheet1.Location = new System.Drawing.Point(30, 40);
            this.comboBoxSheet1.Name = "comboBoxSheet1";
            this.comboBoxSheet1.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSheet1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Source Sheet";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(200, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Destination Sheet";
            // 
            // comboBoxSheet2
            // 
            this.comboBoxSheet2.FormattingEnabled = true;
            this.comboBoxSheet2.Location = new System.Drawing.Point(200, 40);
            this.comboBoxSheet2.Name = "comboBoxSheet2";
            this.comboBoxSheet2.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSheet2.TabIndex = 3;
            // 
            // listBoxSheet1
            // 
            this.listBoxSheet1.FormattingEnabled = true;
            this.listBoxSheet1.Location = new System.Drawing.Point(30, 100);
            this.listBoxSheet1.Name = "listBoxSheet1";
            this.listBoxSheet1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxSheet1.Size = new System.Drawing.Size(120, 30);
            this.listBoxSheet1.TabIndex = 4;
            this.listBoxSheet1.Click += new System.EventHandler(this.listBoxSheet1_Click);
            // 
            // listBoxSheet2
            // 
            this.listBoxSheet2.FormattingEnabled = true;
            this.listBoxSheet2.Location = new System.Drawing.Point(200, 100);
            this.listBoxSheet2.Name = "listBoxSheet2";
            this.listBoxSheet2.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxSheet2.Size = new System.Drawing.Size(120, 30);
            this.listBoxSheet2.TabIndex = 5;
            this.listBoxSheet2.Click += new System.EventHandler(this.listBoxSheet2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(30, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(102, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Source Ref. Column";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(200, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(121, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Destination Ref. Column";
            // 
            // CompareForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(385, 261);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.listBoxSheet2);
            this.Controls.Add(this.listBoxSheet1);
            this.Controls.Add(this.comboBoxSheet2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxSheet1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CompareForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Compare Worksheets";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxSheet1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBoxSheet2;
        private System.Windows.Forms.ListBox listBoxSheet1;
        private System.Windows.Forms.ListBox listBoxSheet2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}