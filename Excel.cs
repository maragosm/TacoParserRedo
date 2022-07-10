using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace TacoTake2
{
    public class Excel
    {
        public string path = Directory.GetCurrentDirectory() + @"\TacoBellLocations.csv";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)ws.Cells[i, j];
            //string cellValue = range.Value.ToString();
            if (range.Value != null)
                return range.Value.ToString();
            else
                return "";
        }

    }
}
