using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Dropbox.Api;

namespace TestAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            DirectoryForm myForm = new DirectoryForm();
            myForm.ShowDialog();
        }

        private void buttonUpper_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges1 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow1 = myRanges1.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            foreach (Excel.Range myRange1 in myRanges1)
            {
                if (myRange1.Column <= lastRow1.Column && myRange1.Row <= lastRow1.Row)
                {
                    string myCellRegion1 = myRange1.Text;
                    string upperCase = myCellRegion1.ToUpper();
                    myRange1.Value2 = upperCase;
                }
                else
                { break; }
            }
        }

        private void buttonLower_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges2 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow2 = myRanges2.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            foreach (Excel.Range myRange2 in myRanges2)
            {
                if (myRange2.Column <= lastRow2.Column && myRange2.Row <= lastRow2.Row)
                {
                    string myCellRegion2 = myRange2.Text;
                    string lowerCase = myCellRegion2.ToLower();
                    myRange2.Value2 = lowerCase;
                }
                else
                { break; }
            }
        }

        private void buttonTitle_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges3 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow3 = myRanges3.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            foreach (Excel.Range myRange3 in myRanges3)
            {
                if (myRange3.Column <= lastRow3.Column && myRange3.Row <= lastRow3.Row)
                {
                    string myCellRegion3 = myRange3.Text;
                    TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                    myRange3.Value2 = myTI.ToTitleCase(myCellRegion3);
                }
                else
                { break; }
            }
        }

        private void buttonCreateFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges4 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow4 = myRanges4.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            string folderPath4 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Directory Path", Title: "Create Directory", Type: 2);

            foreach (Excel.Range myRange4 in myRanges4)
            {
                if (myRange4.Column <= lastRow4.Column && myRange4.Row <= lastRow4.Row)
                {
                    string myCellRegion4 = myRange4.Text;
                    System.IO.Directory.CreateDirectory(folderPath4 + @"\" + myCellRegion4);
                }
                else
                { break; }
            }
            MessageBox.Show("Completed");
        }

        private void buttonCopyFiles_Click(object sender, RibbonControlEventArgs e)
        {
            string sourcePath5 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Source Folder Path", Title: "Source Directory", Type: 2);
            Excel.Range folderPath5 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing Destination Folder Path", Title: "Destination Directory", Type: 8) as Excel.Range;
            Excel.Range lastRow5 = folderPath5.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //Excel.Range filePaths5 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing File Names", Title: "Filename", Type: 8);
            //DialogResult result5 = MessageBox.Show("Do you want to move the files" + Environment.NewLine + "Press No to copy the files", "Alert", MessageBoxButtons.YesNoCancel);
            string[] sourceFiles5 = System.IO.Directory.GetFiles(sourcePath5, ".", System.IO.SearchOption.AllDirectories);
            ProgressUpdater pBar1 = new ProgressUpdater();
            pBar1.Visible = true;
            pBar1.progressBar1.Minimum = 1;
            pBar1.progressBar1.Maximum = lastRow5.Row;
            pBar1.progressBar1.Value = 1;
            pBar1.progressBar1.Step = 1;

            foreach (Excel.Range folderPath in folderPath5)
            {
                //if (folderPath.Column == lastRow5.Column && folderPath.Row <= lastRow5.Row)
                //{
                Excel.Range file5 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 1];
                Excel.Range ext5 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 2];
                string folder5 = Convert.ToString(folderPath.Value2);
                string fileName5 = Convert.ToString(file5.Value2);
                string extension5 = Convert.ToString(ext5.Value2);
                string sourceFile5 = Array.Find(sourceFiles5, s => s.EndsWith(fileName5 + @"." + extension5));
                //if (result5 == DialogResult.Yes)
                //{
                //    try
                //    {
                //    System.IO.File.Move(sourceFile5, System.IO.Path.Combine(folder5,fileName5 + @"." + extension5));
                //    }
                //    catch (Exception e1)
                //    {
                //        Console.Error.WriteLine("The process failed: {0}", e1);
                //    }
                //}
                //if (result5 == DialogResult.No)
                //{
                try
                {
                    System.IO.File.Copy(sourceFile5, System.IO.Path.Combine(folder5, fileName5 + @"." + extension5));
                    pBar1.progressBar1.PerformStep();
                    pBar1.progressBar1.Refresh();
                }
                catch (Exception e2)
                {
                    Console.Error.WriteLine("The process failed: {0}", e2);
                }
                //}
                //if(result5 == DialogResult.Cancel)
                //{
                //break;
                //}
                //}
                //else
                //{ break; }
            }
            pBar1.Dispose();
            MessageBox.Show("Completed");
        }

        private void buttonMoveFiles_Click(object sender, RibbonControlEventArgs e)
        {
            string sourcePath6 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Source Folder Path", Title: "Source Directory", Type: 2);
            Excel.Range folderPath6 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing Destination Folder Path", Title: "Destination Directory", Type: 8) as Excel.Range;

            string[] sourceFiles6 = System.IO.Directory.GetFiles(sourcePath6, ".", System.IO.SearchOption.AllDirectories);

            ProgressUpdater pBar = new ProgressUpdater();
            pBar.Visible = true;
            pBar.progressBar1.Minimum = 1;
            pBar.progressBar1.Maximum = folderPath6.Rows.Count;
            pBar.progressBar1.Value = 1;
            pBar.progressBar1.Step = 1;


            foreach (Excel.Range folderPath in folderPath6)
            {
                Excel.Range file6 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 1];
                Excel.Range ext6 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 2];
                string folder6 = Convert.ToString(folderPath.Value2);
                string fileName6 = Convert.ToString(file6.Value2);
                string extension6 = Convert.ToString(ext6.Value2);
                string sourceFile6 = Array.Find(sourceFiles6, s => s.EndsWith(fileName6 + @"." + extension6));

                try
                {
                    System.IO.File.Move(sourceFile6, System.IO.Path.Combine(folder6, fileName6 + @"." + extension6));
                    pBar.progressBar1.PerformStep();
                    pBar.progressBar1.Refresh();

                }
                catch (Exception e1)
                {
                    Console.Error.WriteLine("The process failed: {0}", e1);
                }
            }
            pBar.Dispose();
            MessageBox.Show("Completed");
        }

        private void buttonDeleteFiles_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range folderPath7 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing Destination Folder Path", Title: "Destination Directory", Type: 8) as Excel.Range;
            DialogResult result7 = MessageBox.Show("Do you really want to delete the files" + Environment.NewLine + "Press No to end the process", "Alert", MessageBoxButtons.YesNo);

            if (result7 == DialogResult.Yes)
            {
                foreach (Excel.Range folderPath in folderPath7)
                {
                    Excel.Range file7 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 1];
                    Excel.Range ext7 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 2];
                    string folder7 = Convert.ToString(folderPath.Value2);
                    string fileName7 = Convert.ToString(file7.Value2);
                    string extension7 = Convert.ToString(ext7.Value2);

                    try
                    {
                        System.IO.File.Delete(System.IO.Path.Combine(folder7, fileName7 + @"." + extension7));
                    }
                    catch (Exception e3)
                    {
                        Console.Error.WriteLine("The process failed: {0}", e3);
                    }
                }
                MessageBox.Show("Completed");
            }
        }

        private void buttonReplaceFiles_Click(object sender, RibbonControlEventArgs e)
        {
            string sourcePath8 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Source Folder Path", Title: "Source Directory", Type: 2);
            Excel.Range folderPath8 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing Destination Folder Path", Title: "Destination Directory", Type: 8) as Excel.Range;
            string[] sourceFiles8 = System.IO.Directory.GetFiles(sourcePath8, ".", System.IO.SearchOption.AllDirectories);
            string backupPath8 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter BackUp Folder Path", Title: "Backup Directory", Type: 2);

            ProgressUpdater pBar = new ProgressUpdater();
            pBar.Visible = true;
            pBar.progressBar1.Minimum = 1;
            pBar.progressBar1.Maximum = folderPath8.Rows.Count;
            pBar.progressBar1.Value = 1;
            pBar.progressBar1.Step = 1;


            foreach (Excel.Range folderPath in folderPath8)
            {
                Excel.Range file8 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 1];
                Excel.Range ext8 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 2];
                string folder8 = Convert.ToString(folderPath.Value2);
                string fileName8 = Convert.ToString(file8.Value2);
                string extension8 = Convert.ToString(ext8.Value2);
                string sourceFile8 = Array.Find(sourceFiles8, s => s.EndsWith(fileName8 + @"." + extension8));

                try
                {
                    System.IO.File.Replace(sourceFile8, System.IO.Path.Combine(folder8, fileName8 + @"." + extension8), System.IO.Path.Combine(backupPath8, fileName8 + @"." + extension8), ignoreMetadataErrors: true);
                    pBar.progressBar1.PerformStep();
                    pBar.progressBar1.Refresh();

                }
                catch (Exception e4)
                {
                    Console.Error.WriteLine("The process failed: {0}", e4);
                }
            }
            pBar.Dispose();
            MessageBox.Show("Completed");
        }

        private void buttonDeleteDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges9 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow9 = myRanges9.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            string folderPath9 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Directory Path", Title: "Delete Directory", Type: 2);
            DialogResult result8 = MessageBox.Show("Do you really want to delete the folders" + Environment.NewLine + "This process cannot be undone!" + Environment.NewLine + "Press No to end the process", "Alert", MessageBoxButtons.YesNo);

            if (result8 == DialogResult.Yes)
            {
                foreach (Excel.Range myRange9 in myRanges9)
                {
                    if (myRange9.Column <= lastRow9.Column && myRange9.Row <= lastRow9.Row)
                    {
                        string myCellRegion4 = myRange9.Text;
                        System.IO.Directory.Delete(folderPath9 + @"\" + myCellRegion4, recursive: true);
                    }
                    else
                    { break; }
                }
                MessageBox.Show("Completed");
            }
        }

        private void buttonDeleteEmptyDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            string folderPath10 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Directory Path", Title: "Delete Empty Directory", Type: 2);
            deleteEmptyDirs(folderPath10);
            MessageBox.Show("Empty Folders Deleted");
        }

        private void deleteEmptyDirs(string folderPath10)
        {
            //List<string> elements10 = new List<string>();
            if (String.IsNullOrEmpty(folderPath10))
            { MessageBox.Show("Please enter all references"); }
            //{ throw new ArgumentException("Please enter the reference"); }
            else
            {
                try
                {
                    foreach (var folder10 in System.IO.Directory.EnumerateDirectories(folderPath10))
                    {
                        deleteEmptyDirs(folder10);
                    }
                    var entries10 = System.IO.Directory.EnumerateFileSystemEntries(folderPath10);
                    if (!entries10.Any())
                    {
                        try
                        {
                            System.IO.Directory.Delete(folderPath10);
                            //elements10.Add("string");
                        }
                        catch (UnauthorizedAccessException) { }
                        catch (System.IO.DirectoryNotFoundException) { }
                    }
                }
                catch (UnauthorizedAccessException) { }
            }
        }

        private void buttonTrimSpace_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges12 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow12 = myRanges12.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            char[] charsToTrim = { ' ', '\n' };

            foreach (Excel.Range myRange12 in myRanges12)
            {
                if (myRange12.Column <= lastRow12.Column && myRange12.Row <= lastRow12.Row)
                {
                    string myCellRegion12 = myRange12.Text;
                    string cellTrim = myCellRegion12.Trim(trimChars: charsToTrim);
                    string spaceTrim = System.Text.RegularExpressions.Regex.Replace(cellTrim, @"\s+", " ");
                    myRange12.Value2 = spaceTrim;
                }
                else
                { break; }
            }
        }

        private void buttonUnhideFiles_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range folderPath14 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing Complete File Path", Title: "Destination Files", Type: 8) as Excel.Range;

            foreach (Excel.Range folderPath in folderPath14)
            {
                Excel.Range file14 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column];
                //Excel.Range ext14 = Globals.ThisAddIn.Application.Cells[folderPath.Row, folderPath.Column + 2];
                //string folder14 = folderPath.Value2;
                string fileName14 = file14.Value2;
                //string extension14 = ext14.Value2;

                try
                {
                    System.IO.File.SetAttributes(fileName14, fileAttributes: System.IO.FileAttributes.Normal);

                }
                catch (Exception e14)
                {
                    Console.Error.WriteLine("The process failed: {0}", e14);
                }
            }
            MessageBox.Show("Completed");
        }

        private void buttonJoin_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges13 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string concValue = Globals.ThisAddIn.Application.InputBox("Enter the delimiter", "Concatenate Values", Type: 2);
            Excel.Range lastCell = myRanges13.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //resultCell.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            Excel.Range wrkngRange = Globals.ThisAddIn.Application.Cells[1,myRanges13.Column + myRanges13.Columns.Count];
            wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                for (double j = 1; j <= myRanges13.Rows.Count; j++)
                {
                    if (j <= lastCell.Row)
                    {
                        string strword = "";
                        for (double i = 1; i <= myRanges13.Columns.Count; i++)
                        {
                            string cellValue = myRanges13.Cells[j, i].text;
                            if (!strword.Any())
                            { strword = cellValue; }
                            else
                            { strword = strword + concValue + cellValue; }
                        }
                        string resultValue = strword;
                        myRanges13.Cells[j, myRanges13.Columns.Count + 1].Value2 = resultValue;
                    }
                    else
                    { break; }
                }
            MessageBox.Show("Completed");
        }

        private void buttonTrimEnter_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges12 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow12 = myRanges12.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            char[] charsToTrim = { ' ', '\n' };

            foreach (Excel.Range myRange12 in myRanges12)
            {
                if (myRange12.Column <= lastRow12.Column && myRange12.Row <= lastRow12.Row)
                {
                    string myCellRegion12 = myRange12.Text;
                    string cellTrim = myCellRegion12.Trim(trimChars: charsToTrim);
                    string enterTrim = System.Text.RegularExpressions.Regex.Replace(cellTrim, @"\n+", "\n");
                    string frontSpace = System.Text.RegularExpressions.Regex.Replace(enterTrim, @" *\n", "\n");
                    string endSpace = System.Text.RegularExpressions.Regex.Replace(frontSpace, @"\n *", "\n");
                    string spaceTrim = System.Text.RegularExpressions.Regex.Replace(endSpace, @" +", " ");
                    myRange12.Value2 = spaceTrim;
                }
                else
                { break; }
            }
        }

        private void buttonCombineExcel_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog browseExcel = new OpenFileDialog();
            browseExcel.Multiselect = true;
            browseExcel.Filter = "Excel Files (.xls)|*.xls*";
            browseExcel.FilterIndex = 1;
            DialogResult result = browseExcel.ShowDialog();
            string[] sourceFilePaths = browseExcel.FileNames;
            Excel.Workbook destFile = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (result == DialogResult.OK)
            {
                try
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = false;
                    foreach (string sourceFilePath in sourceFilePaths)
                    {
                        Excel.Workbook sourceFile = Globals.ThisAddIn.Application.Workbooks.Open(sourceFilePath);
                        int wsIndx = sourceFile.Worksheets.Count;
                        for (int i = 1; i <= wsIndx; i++)
                        {
                            Excel.Worksheet sourceWs = sourceFile.Worksheets[i];
                            sourceWs.Copy(After: destFile.Worksheets[destFile.Worksheets.Count]);
                        }
                        sourceFile.Close(SaveChanges: false);
                    }
                    MessageBox.Show("Completed");
                }
                catch (Exception e15)
                { Console.Error.WriteLine("The process failed : {0}", e15); }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void buttonMoveDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range myRanges11 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow11 = myRanges11.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            string sourcePath11 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Source Folder Path", Title: "Source Directory", Type: 2);
            string folderPath11 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Destination Folder Path", Title: "Destination Directory", Type: 2);

            ProgressUpdater pBar = new ProgressUpdater();
            pBar.Visible = true;
            pBar.progressBar1.Minimum = 1;
            pBar.progressBar1.Maximum = lastRow11.Row;
            pBar.progressBar1.Value = 1;
            pBar.progressBar1.Step = 1;

            try
            {
                foreach (Excel.Range myRange11 in myRanges11)
                {
                    if (myRange11.Column <= lastRow11.Column && myRange11.Row <= lastRow11.Row)
                    {
                        string myCellRegion11 = myRange11.Text;
                        string destPath11 = folderPath11.ToString();
                        string srcPath11 = sourcePath11.ToString();
                        System.IO.Directory.Move(System.IO.Path.Combine(srcPath11, myCellRegion11), System.IO.Path.Combine(destPath11, myCellRegion11));
                        pBar.progressBar1.PerformStep();
                        pBar.progressBar1.Refresh();
                    }
                }
            }
            catch (Exception e6)
            {
                Console.Error.WriteLine("The process failed: {0}", e6);
            }
            pBar.Dispose();
            MessageBox.Show("Completed");
        }

        private void buttonCombineSheets_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook destWb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet destWs = destWb.Worksheets.Add(Before: destWb.Worksheets[1]);
            int wsCount = destWb.Worksheets.Count;

            for (int i = 2; i <= wsCount; i++)
            {
                Excel.Range wrkngRange = destWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastRow = wrkngRange.Row;
                Excel.Worksheet sourceWs = destWb.Worksheets[i];
                Excel.Range sourceRange = sourceWs.UsedRange;
                sourceRange.Copy(Destination: destWs.Cells[lastRow + 1, 1]);
            }
            MessageBox.Show("Completed");
        }

        private void buttonSplitSheets_Click(object sender, RibbonControlEventArgs e)
        {
            var noOfCell = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter the number of rows to Split", Title: "Split Worksheet", Type: 1);
            int noOfCells = Convert.ToInt32(noOfCell);
            Excel.Workbook currentWb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sourceWs = currentWb.ActiveSheet;
            Excel.Range wrkRange = sourceWs.UsedRange;
            int noOfRow = wrkRange.Rows.Count;
            int noOfColumn = wrkRange.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int wrkRow = (noOfRow / noOfCells) + 1;

            for (int i = 1; i <= wrkRow; i++)
            {
                int firstCell = (noOfCells * (i - 1)) + 1;
                int lastCell = (noOfCells * i);
                Excel.Worksheet destWs = currentWb.Worksheets.Add(After: currentWb.Worksheets[currentWb.Worksheets.Count]);
                sourceWs.Range[sourceWs.Cells[firstCell, 1], sourceWs.Cells[lastCell, noOfColumn]].Copy(Destination: destWs.Range["A2"]);
            }
            MessageBox.Show("Completed");
        }

        private void buttonLeftNth_Click(object sender, RibbonControlEventArgs e)
        {
            string delimChar = Globals.ThisAddIn.Application.InputBox("Enter the delimiter", "Split at Nth Occurence");
            double nthIndex = Globals.ThisAddIn.Application.InputBox("Enter the Nth index of the delimiter", "Split at Nth Occurence", Type: 1);
            int n = Convert.ToInt32(nthIndex);
            char t = Convert.ToChar(delimChar);
            Excel.Range myRanges16 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow16 = myRanges16.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int j = 0;
            Excel.Range wrkngRange = Globals.ThisAddIn.Application.Cells[1, myRanges16.Column + myRanges16.Columns.Count];
            wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            foreach (Excel.Range myRange16 in myRanges16)
                {
                    if (myRange16.Column <= lastRow16.Column && myRange16.Row <= lastRow16.Row)
                    {
                        string myCellRegion12 = myRange16.Text;
                        int strLen = myCellRegion12.Length;
                        int count = 0;
                        for (int i = 0; i < strLen; i++)
                        {
                            if (myCellRegion12[i] == t)
                            {
                                count++;
                                if (count == n)
                                {
                                    j = i;
                                }
                            }
                        }
                    if (j != 0 && strLen!= 0)
                    {
                        myRange16.Offset[0, 1].Value2 = myCellRegion12.Substring(0, j);
                        myRange16.Offset[0, 2].Value2 = myCellRegion12.Substring(j + 1, strLen - (j + 1));
                    }
                    else if (strLen == 0 && myRange16.Row < lastRow16.Row)
                    { continue; }
                    else
                    {
                        wrkngRange.Offset[0, -1].EntireColumn.Delete();
                        wrkngRange.Offset[0, -1].EntireColumn.Delete();
                        MessageBox.Show("The delimiter do not exist within the selected range.");
                        return;
                    }
                }
                    else
                    { break; }
                }
                MessageBox.Show("Completed");
        }

        private void buttonMidNth_Click(object sender, RibbonControlEventArgs e)
        {
            string delimChar = Globals.ThisAddIn.Application.InputBox("Enter the delimiter", "Split at Nth Occurence");
            double mthIndex = Globals.ThisAddIn.Application.InputBox("Enter the first index of the delimiter", "Split at Nth Occurence", Type: 1);
            double nthIndex = Globals.ThisAddIn.Application.InputBox("Enter the last index of the delimiter", "Split at Nth Occurence", Type: 1);
            int m = Convert.ToInt32(mthIndex);
            int n = Convert.ToInt32(nthIndex);
            char t = Convert.ToChar(delimChar);
            Excel.Range myRanges16 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow16 = myRanges16.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int j = 0;
            int z = 0;

            Excel.Range wrkngRange = Globals.ThisAddIn.Application.Cells[1, myRanges16.Column + myRanges16.Columns.Count];
            wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                foreach (Excel.Range myRange16 in myRanges16)
                {
                    if (myRange16.Column <= lastRow16.Column && myRange16.Row <= lastRow16.Row)
                    {
                        string myCellRegion12 = myRange16.Text;
                        int strLen = myCellRegion12.Length;
                        int count = 0;
                        for (int i = 0; i < strLen; i++)
                        {
                            if (myCellRegion12[i] == t)
                            {
                                count++;
                                if (count == m)
                                {
                                    j = i;
                                }
                                else if (count == n)
                                {
                                    z = i;
                                }
                            }
                        }
                        if (j != 0 && z != 0 && strLen != 0)
                        {
                            myRange16.Offset[0, 1].Value2 = myCellRegion12.Substring(j + 1, z - j - 1);
                        }
                        else if (strLen == 0 && myRange16.Row < lastRow16.Row)
                        { continue; }
                        else
                        {
                        wrkngRange.Offset[0,-1].EntireColumn.Delete();
                        MessageBox.Show("The delimiter do not exist within the selected range.");
                        return;
                        }
                    }
                    else
                    { break; }
                }
                MessageBox.Show("Completed");
        }

        private void buttonRightNth_Click(object sender, RibbonControlEventArgs e)
        {
            string delimChar = Globals.ThisAddIn.Application.InputBox("Enter the delimiter", "Split at Nth Occurence");
            double nthIndex = Globals.ThisAddIn.Application.InputBox("Enter the Nth index of the delimiter", "Split at Nth Occurence", Type: 1);
            int n = Convert.ToInt32(nthIndex);
            char t = Convert.ToChar(delimChar);
            Excel.Range myRanges16 = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Range lastRow16 = myRanges16.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int j = 0;
            Excel.Range wrkngRange = Globals.ThisAddIn.Application.Cells[1, myRanges16.Column + myRanges16.Columns.Count];
            wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            wrkngRange.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                foreach (Excel.Range myRange16 in myRanges16)
                {
                    if (myRange16.Column <= lastRow16.Column && myRange16.Row <= lastRow16.Row)
                    {
                        string myCellRegion12 = myRange16.Text;
                        int strLen = myCellRegion12.Length;
                        int count = 0;
                        for (int i = strLen - 1; i > 0; i--)
                        {
                            if (myCellRegion12[i] == t)
                            {
                                count++;
                                if (count == n)
                                {
                                    j = i;
                                }
                            }
                        }
                    if (j != 0 && strLen != 0)
                    {
                        myRange16.Offset[0, 1].Value2 = myCellRegion12.Substring(0, j);
                        myRange16.Offset[0, 2].Value2 = myCellRegion12.Substring(j + 1, strLen - (j + 1));
                    }
                    else if (strLen == 0 && myRange16.Row < lastRow16.Row)
                    { continue; }
                    else
                    {
                        wrkngRange.Offset[0, -1].EntireColumn.Delete();
                        wrkngRange.Offset[0, -1].EntireColumn.Delete();
                        MessageBox.Show("The delimiter do not exist within the selected range.");
                        return;
                    }
                }
                    else
                    { break; }
                }
                MessageBox.Show("Completed");
        }

        private void buttonHyperlink_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook srcWbk = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet srcWrst = srcWbk.Worksheets.Add(Before: srcWbk.Worksheets[1]);
            srcWrst.Name = "Hyperlink";

            for (int i = 2; i <= srcWbk.Worksheets.Count; i++)
            {
                srcWrst.Hyperlinks.Add(Anchor: srcWrst.Range["A" + i], Address:"", SubAddress: srcWbk.Worksheets[i].Name + "!A1",
                    ScreenTip: "Go to " + srcWbk.Worksheets[i].Name, TextToDisplay:srcWbk.Worksheets[i].Name);
            }
            srcWrst.Activate();
        }

        private void buttonRenameFiles_Click(object sender, RibbonControlEventArgs e)
        {
            string sourcePath16 = Globals.ThisAddIn.Application.InputBox(Prompt: "Enter Source Folder Path", Title: "Source Directory", Type: 2);
            Excel.Range oldFileName16 = Globals.ThisAddIn.Application.InputBox(Prompt: "Select Cell Containing Old File Name", Title: "Source Directory", Type: 8) as Excel.Range;

            string[] sourceFiles16 = System.IO.Directory.GetFiles(sourcePath16, ".", System.IO.SearchOption.AllDirectories);

            ProgressUpdater pBar = new ProgressUpdater();
            pBar.Visible = true;
            pBar.progressBar1.Minimum = 1;
            pBar.progressBar1.Maximum = oldFileName16.Rows.Count;
            pBar.progressBar1.Value = 1;
            pBar.progressBar1.Step = 1;

            foreach (Excel.Range filePath in oldFileName16)
            {
                Excel.Range newFileName16 = Globals.ThisAddIn.Application.Cells[filePath.Row, filePath.Column + 1];
                Excel.Range ext16 = Globals.ThisAddIn.Application.Cells[filePath.Row, filePath.Column + 2];
                string oldName16 = Convert.ToString(filePath.Value2);
                string fileName16 = Convert.ToString(newFileName16.Value2);
                string extension16 = Convert.ToString(ext16.Value2);
                string sourceFile16 = Array.Find(sourceFiles16, s => s.EndsWith(oldName16 + @"." + extension16));

                try
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(sourceFile16, fileName16 + "." + extension16);
                    pBar.progressBar1.PerformStep();
                    pBar.progressBar1.Refresh();
                }
                catch (Exception e1)
                {
                    Console.Error.WriteLine("The process failed: {0}", e1);
                }
            }
            pBar.Dispose();
            MessageBox.Show("Completed");
        }

        private async void buttonFileList_Click(object sender, RibbonControlEventArgs e)
        {
            string accesstoken = Globals.ThisAddIn.Application.InputBox("Please Enter Your Access Token");

            using (var dbx = new DropboxClient(accesstoken))
            {
                var full = await dbx.Users.GetCurrentAccountAsync();
                DialogResult userName = MessageBox.Show("Press 'Yes' to list the files in this dropbox account." + System.Environment.NewLine +
                    System.Environment.NewLine + "Display Name: " + full.Name.DisplayName + System.Environment.NewLine + "User Email: " + full.Email,
                    "Dropbox Account Information", MessageBoxButtons.YesNo);

                if (userName == DialogResult.Yes)
                {
                    var listFiles = await dbx.Files.ListFolderAsync(string.Empty, recursive: true);
                    int j = 2;

                    Globals.ThisAddIn.Application.Cells[1, 1].Value2 = "File Path";
                    Globals.ThisAddIn.Application.Cells[1, 2].Value2 = "File Name";
                    Globals.ThisAddIn.Application.Cells[1, 3].Value2 = "File ID";
                    Globals.ThisAddIn.Application.Cells[1, 4].Value2 = "Rev";
                    Globals.ThisAddIn.Application.Cells[1, 5].Value2 = "Size (in Bytes)";
                    Globals.ThisAddIn.Application.Cells[1, 6].Value2 = "Client Modified Date";
                    Globals.ThisAddIn.Application.Columns[6].NumberFormat = "dd/mm/yyyy";
                    Globals.ThisAddIn.Application.Cells[1, 7].Value2 = "Client Modified Time";
                    Globals.ThisAddIn.Application.Columns[7].NumberFormat = "hh:mm:ss";
                    Globals.ThisAddIn.Application.Cells[1, 8].Value2 = "Server Modified Date";
                    Globals.ThisAddIn.Application.Columns[8].NumberFormat = "dd/mm/yyyy";
                    Globals.ThisAddIn.Application.Cells[1, 9].Value2 = "Server Modified Time";
                    Globals.ThisAddIn.Application.Columns[9].NumberFormat = "hh:mm:ss";

                    foreach (var fileList in listFiles.Entries.Where(i => i.IsFile))
                    {
                        Globals.ThisAddIn.Application.Cells[j, 1].Value2 = fileList.AsFile.PathDisplay;
                        Globals.ThisAddIn.Application.Cells[j, 2].Value2 = fileList.AsFile.Name;
                        Globals.ThisAddIn.Application.Cells[j, 3].Value2 = fileList.AsFile.Id;
                        Globals.ThisAddIn.Application.Cells[j, 4].Value2 = fileList.AsFile.Rev;
                        Globals.ThisAddIn.Application.Cells[j, 5].Value2 = fileList.AsFile.Size;
                        Globals.ThisAddIn.Application.Cells[j, 6].Value2 = fileList.AsFile.ClientModified.ToOADate();
                        Globals.ThisAddIn.Application.Cells[j, 7].Value2 = fileList.AsFile.ClientModified.ToLocalTime();
                        Globals.ThisAddIn.Application.Cells[j, 8].Value2 = fileList.AsFile.ServerModified.ToOADate();
                        Globals.ThisAddIn.Application.Cells[j, 9].Value2 = fileList.AsFile.ServerModified.ToLocalTime();
                        j++;
                    }
                }
                else
                {
                    return;
                }
            }
        }

        //private void buttonCompareSheet_Click(object sender, RibbonControlEventArgs e)
        //{
        //    CompareForm formCompare = new CompareForm();
        //    formCompare.ShowDialog();
        //}
    }
}