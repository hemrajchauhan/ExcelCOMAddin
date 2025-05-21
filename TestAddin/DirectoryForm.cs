using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestAddin
{
    public partial class DirectoryForm : Form
    {
        public DirectoryForm()
        {
            InitializeComponent();
        }

        private void buttonExtension_Click(object sender, EventArgs e)
        {
            {
                if (textBoxPath.Text != "")
                {
                    comboBoxExtension.Items.Clear();
                    string folderPath = textBoxPath.Text;
                    string[] fileNames = Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories);

                    TextWriter fileExt = new StreamWriter(System.IO.Path.GetTempPath() + @"\files_ext.txt");

                    foreach (string fileName in fileNames)
                    {
                        string allFileExt = System.IO.Path.GetExtension(fileName);
                        fileExt.WriteLine(allFileExt);
                    }
                    fileExt.Close();

                    string[] fileContent = File.ReadAllLines(System.IO.Path.GetTempPath() + @"\files_ext.txt");
                    string[] distinctExt = fileContent.Distinct().ToArray();

                    foreach (string uniqueExt in distinctExt)
                    {
                        comboBoxExtension.Items.Add(uniqueExt);
                    }
                    comboBoxExtension.Items.Add("All Files");
                    File.Delete(System.IO.Path.GetTempPath() + @"\files_ext.txt");
                }
                else
                {
                    MessageBox.Show("Please enter Directory Path");
                }
            }
        }

        private void buttonFileList_Click(object sender, EventArgs e)
        {
            {
                if (textBoxPath.Text != "")
                {
                    if (comboBoxExtension.Text == "All Files")
                    {
                        string myExt = "*";
                        string folderPath1 = textBoxPath.Text;
                        string[] fileNames1 = Directory.GetFiles(folderPath1, myExt, SearchOption.AllDirectories);
                        //Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ActiveSheet;
                        //int i = mySheet.Range["A:A"].Rows.Count;
                        //int j = i + 1;
                        int j = 1;
                        Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1",Type.Missing];
                        wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                        foreach (string fileName1 in fileNames1)
                        {
                            Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + j];
                            myRange.Value2 = fileName1;
                            j++;
                        }
                        MessageBox.Show("Completed");
                    }
                    else
                    {
                        string myExt = "*" + comboBoxExtension.Text;
                        string folderPath2 = textBoxPath.Text;
                        string[] fileNames2 = Directory.GetFiles(folderPath2, myExt, SearchOption.AllDirectories);
                        //Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ActiveSheet;
                        //int i = mySheet.Range["A:A"].Rows.Count;
                        //int j = i + 1;
                        int j = 1;
                        Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1", Type.Missing];
                        wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                        foreach (string fileName2 in fileNames2)
                        {
                            Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + j];
                            myRange.Value2 = fileName2;
                            j++;
                        }
                        MessageBox.Show("Completed");
                    }
                }
                else
                {
                    MessageBox.Show("Please enter Directory Path");
                }
            }
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog myFolder = new FolderBrowserDialog();
            DialogResult resultFolder = myFolder.ShowDialog();
            textBoxPath.Text = myFolder.SelectedPath;
        }

        private void DirectoryForm_Load(object sender, EventArgs e)
        {

        }

        private void buttonFileCount_Click(object sender, EventArgs e)
        {
            if (textBoxPath.Text != "")
            {
                string folderPath3 = textBoxPath.Text;
                if (comboBoxExtension.Text == "All Files")
                {
                    string myExt1 = "*";
                    int myFileCount1 = Directory.GetFiles(folderPath3, myExt1, SearchOption.AllDirectories).Length;
                    textBoxFileCount.Text = myFileCount1.ToString();
                }
                else
                {
                    string myExt1 = "*" + comboBoxExtension.Text;
                    int myFileCount2 = Directory.GetFiles(folderPath3, myExt1, SearchOption.AllDirectories).Length;
                    textBoxFileCount.Text = myFileCount2.ToString();
                }
            }
            else
            {
                MessageBox.Show("Please enter Directory Path");
            }
        }

        private void buttonFolderList_Click(object sender, EventArgs e)
        {
            if (textBoxPath.Text != "")
            {
                string folderPath4 = textBoxPath.Text;
                if (comboBoxExtension.Text == "All Files")
                {
                    string myExt3 = "*";
                    string[] folderLists = Directory.GetDirectories(folderPath4, "*", SearchOption.AllDirectories);
                    int z = 1;
                    DialogResult folderList1 = MessageBox.Show("Press Yes to get folder list with counts","Folder List",MessageBoxButtons.YesNo);
                    Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1", Type.Missing];
                    wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    foreach (string folderList in folderLists)
                    {
                        int myFileCount3 = Directory.GetFiles(folderList, myExt3, SearchOption.TopDirectoryOnly).Length;
                        Excel.Range myRange1 = Globals.ThisAddIn.Application.Range["A" + z];
                        if (folderList1 == DialogResult.Yes)
                        {
                            myRange1.Value2 = folderList + " : " + myFileCount3;
                            z++;
                        }
                        else
                        {
                            myRange1.Value2 = folderList;
                            z++;
                        }
                    }
                    MessageBox.Show("Complete");
                }
                else
                {
                    string myExt3 = "*" + comboBoxExtension.Text;
                    string[] folderLists = Directory.GetDirectories(folderPath4, "*", SearchOption.AllDirectories);
                    int z = 1;
                    DialogResult folderList1 = MessageBox.Show("Press Yes to get folder list with counts", "Folder List", MessageBoxButtons.YesNo);
                    Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1", Type.Missing];
                    wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    foreach (string folderList in folderLists)
                    {
                        int myFileCount3 = Directory.GetFiles(folderList, myExt3, SearchOption.TopDirectoryOnly).Length;
                        Excel.Range myRange1 = Globals.ThisAddIn.Application.Range["A" + z];
                        if (folderList1 == DialogResult.Yes)
                        {
                            myRange1.Value2 = folderList + " : " + myFileCount3;
                            z++;
                        }
                        else
                        {
                            myRange1.Value2 = folderList;
                            z++;
                        }
                    }
                    MessageBox.Show("Complete");
                }
            }
            else
            {
                MessageBox.Show("Please enter Directory Path");
            }
        }

        private void button0KbFileList_Click(object sender, EventArgs e)
        {
            {
                if (textBoxPath.Text != "")
                {
                    if (comboBoxExtension.Text == "All Files")
                    {
                        string myExt5 = "*";
                        string folderPath5 = textBoxPath.Text;
                        DirectoryInfo folderName5 = new DirectoryInfo(folderPath5);
                        FileInfo[] fileNames5 = folderName5.GetFiles(myExt5, SearchOption.AllDirectories);
                        int j = 1;
                        Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1", Type.Missing];
                        wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        DialogResult result5 = MessageBox.Show("Press Yes to get 0Kb File List" + Environment.NewLine + "Press No to get hidden File List", "0Kb or Hidden File List", MessageBoxButtons.YesNo);
                        var filtered5 = fileNames5.Where(f => f.Attributes.HasFlag(FileAttributes.Hidden));

                        if (result5 == DialogResult.Yes)
                        {
                            foreach (FileInfo fileName5 in fileNames5)
                            {
                                if (fileName5.Length == 0)
                                {
                                    Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + j];
                                    myRange.Value2 = fileName5.FullName;
                                    j++;
                                }
                            }
                        }
                        else
                        {
                            foreach (var filter5 in filtered5)
                            {
                                Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + j];
                                myRange.Value2 = filter5.FullName;
                                j++;
                            }
                        }
                            MessageBox.Show("Completed");
                    }
                    else
                    {
                        string myExt5 = "*" + comboBoxExtension.Text;
                        string folderPath5 = textBoxPath.Text;
                        DirectoryInfo folderName5 = new DirectoryInfo(folderPath5);
                        FileInfo[] fileNames5 = folderName5.GetFiles(myExt5, SearchOption.AllDirectories);
                        int j = 1;
                        Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1", Type.Missing];
                        wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        DialogResult result5 = MessageBox.Show("Press Yes to get 0Kb File List" + Environment.NewLine + "Press No to get hidden File List", "0Kb or Hidden File List", MessageBoxButtons.YesNo);
                        var filtered5 = fileNames5.Where(f => f.Attributes.HasFlag(FileAttributes.Hidden));

                        if (result5 == DialogResult.Yes)
                        {
                            foreach (FileInfo fileName5 in fileNames5)
                            {
                                if (fileName5.Length == 0)
                                {
                                    Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + j];
                                    myRange.Value2 = fileName5.FullName;
                                    j++;
                                }
                            }
                        }
                        else
                        {
                            foreach (var filter5 in filtered5)
                            {
                                Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + j];
                                myRange.Value2 = filter5.FullName;
                                j++;
                            }
                        }
                        MessageBox.Show("Completed");
                    }
                }
                else
                {
                    MessageBox.Show("Please enter Directory Path");
                }
            }
        }

        private void buttonEmptyDirectoryList_Click(object sender, EventArgs e)
        {
            if (textBoxPath.Text != "")
            {
                string folderPath6 = textBoxPath.Text;
                Excel.Range wrkngRange = Globals.ThisAddIn.Application.Range["A1", Type.Missing];
                wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                try
                {
                    int k = 1;
                    foreach (var folder6 in System.IO.Directory.EnumerateDirectories(folderPath6, ".", SearchOption.AllDirectories))
                    {
                        var entries10 = System.IO.Directory.EnumerateFileSystemEntries(folder6);
                        if (!entries10.Any())
                        {
                            try
                            {
                                Excel.Range myRange = Globals.ThisAddIn.Application.Range["A" + k];
                                myRange.Value2 = folder6;
                                k++;
                            }
                            catch (UnauthorizedAccessException) { }
                            catch (System.IO.DirectoryNotFoundException) { }
                        }
                    }
                }
                catch (UnauthorizedAccessException) { }
                MessageBox.Show("Completed");
            }
            else
            { MessageBox.Show("Please enter Directory Path"); }
        }
    }
}
