using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestAddin
{
    public partial class CompareForm : Form
    {
        public CompareForm()
        {
            InitializeComponent();
            Excel.Workbook activeWrkbk = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
            int sheetCnt = activeWrkbk.Sheets.Count;
            
            for(int i = 1; i <= sheetCnt; i++)
            {
                comboBoxSheet1.Items.Add(activeWrkbk.Worksheets[i].Name);
                comboBoxSheet2.Items.Add(activeWrkbk.Worksheets[i].Name);
            }
        }

        private void listBoxSheet1_Click(object sender, EventArgs e)
        {
            if (comboBoxSheet1.Text != null)
            {
                Excel.Workbook activeWrkbk = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
                Excel.Worksheet sheet1 = activeWrkbk.Worksheets[comboBoxSheet1.Text];
                Excel.Range wrkRange = sheet1.UsedRange;
                Excel.Range lastCell = sheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int firstColumn = wrkRange.Cells[1, 1].Column;
                int lastColumn = lastCell.Column;

                for (int i = firstColumn; i <= lastColumn; i++)
                {
                    if (sheet1.Cells[1, i].Value2 != null)
                    { listBoxSheet1.Items.Add(sheet1.Cells[1, i].Value2); }
                    else
                    { listBoxSheet1.Items.Add("Column " + i.ToString()); }
                }
            }

        }

        private void listBoxSheet2_Click(object sender, EventArgs e)
        {
            if (comboBoxSheet2.Text != null)
            {
                Excel.Workbook activeWrkbk = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
                Excel.Worksheet sheet2 = activeWrkbk.Worksheets[comboBoxSheet2.Text];
                Excel.Range wrkRange = sheet2.UsedRange;
                Excel.Range lastCell = sheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int firstColumn = wrkRange.Cells[1, 1].Column;
                int lastColumn = lastCell.Column;

                for (int i = firstColumn; i <= lastColumn; i++)
                {
                    if (sheet2.Cells[1, i].Value2 != null)
                    { listBoxSheet2.Items.Add(sheet2.Cells[1, i].Value2); }
                    else
                    { listBoxSheet2.Items.Add("Column " + i.ToString()); }
                }
            }
        }

    }
}
