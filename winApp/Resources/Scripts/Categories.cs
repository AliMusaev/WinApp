using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace winApp.Resources
{
    class Categories
    {
        Excel.Workbook book;
        public Workbook Book { get => book; private set => book = value; }
        Excel.Application excelApp;
        
        public List<string> sheets;

        

        public List<string> LoadCategories()
        {
            sheets = new List<string>();
            string path = Directory.GetCurrentDirectory() + "\\File.xlsx";
            excelApp = new Excel.Application();
            Book = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing);
            for (int i = 1; i <= Book.Sheets.Count; i++)
            {
                sheets.Add(Book.Sheets[i].Name);
            }
            return sheets;
        }
        public void Close()
        {
            Book.Close();
            excelApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            
            
        }

        
    }
}

