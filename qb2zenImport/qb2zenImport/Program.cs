using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Excel;

namespace qb2zenImport
{
    class ImportToZencart
    {
        static void Main(string[] args)
        {
            Excel.Workbook theWorkbook = Excel.Workbooks.Open(
         openFileDialog1.FileName, 0, true, 5,
          "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
          0, true);
            Excel.Sheets sheets = theWorkbook.Worksheets;
            Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
            for (int i = 1; i <= 10; i++)
            {
                Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());
                System.Array myvalues = (System.Array)range.Cells.Value;
                string[] strArray = ConvertToStringArray(myvalues);
            }
        }
    }
}
