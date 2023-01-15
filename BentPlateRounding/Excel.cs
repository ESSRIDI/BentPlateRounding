using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace BentPlateRounding
{
    public class Excel
    {
        string path = "";
        _Application excelApp = new _Excel.Application();
        Workbook wB;
        Worksheet wS;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wB = excelApp.Workbooks.Open(path);
            wS = wB.Worksheets[Sheet];
        }
        public string ExcelReaderCell(int i, int j)
        {
            i++;
            j++;

            if (wS.Cells[i, j].Value2 != null)
            {
                return wS.Cells[i, j].Value2;
            }
            else return "";
        }


    }
}
