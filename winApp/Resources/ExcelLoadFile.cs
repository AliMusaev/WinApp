using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace winApp.Resources
{
    class ExcelLoadFile
    {
        private Excel.Workbook book;
        public List<string> sheets = new List<string>();

        public void Open ()
        {
            Excel.Application excelApp = new Excel.Application();
            book = excelApp.Workbooks.Open("Z:\\Games\\winApp\\File.xlsx", Type.Missing, Type.Missing, Type.Missing, 
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                                           Type.Missing, Type.Missing, Type.Missing);
            for (int i = 1; i < book.Sheets.Count; i++)
            {
                sheets.Add(book.Sheets[i].Name);
            }
        }
    }
}
