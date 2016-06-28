using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelDataPopulateToCombo
{
    public partial class Form1 : Form
    {
        DataTable table = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ReadExcelAndBuildDataTable();
            comboBox1.DataSource = table;
            comboBox1.DisplayMember = "Stud Bolt Size";
            comboBox1.ValueMember = "Value";
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void ReadExcelAndBuildDataTable()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\D\Work\Project\SpeachReconition\ReadExcelDataPopulateToCombo\ReadExcelDataPopulateToCombo\Data\Parameter_Final_V1.1.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.get_Range("J9:T9", "J9:J14");

            //Col
            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                table.Columns.Add(Convert.ToString((range.Cells[1, cCnt] as Excel.Range).Value2));
            }
            //Rows
            for (int row = 2; row < range.Rows.Count + 1; row++)
            {
                DataRow dataRow = table.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {

                    dataRow[col - 1] = Convert.ToString((range.Cells[row, col] as Excel.Range).Value2);
                }
                table.Rows.Add(dataRow);

            }
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
