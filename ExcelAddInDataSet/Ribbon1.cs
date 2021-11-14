using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInDataSet
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
//show form1

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //get notes for column G
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            var activeApp = Globals.ThisAddIn.GetActiveApp();

            int sRowCount = currentSheet.UsedRange.Rows.Count;
            Range sNum = currentSheet.Range["A2:A" + sRowCount];


        }
    }
}
