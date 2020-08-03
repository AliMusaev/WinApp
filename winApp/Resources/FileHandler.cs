using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace winApp.Resources
{
    class FileHandler
    {

        private Application excelApp = null;
        private Workbooks books = null;
        private Workbook book = null;
        private Sheets sheets = null;
        private Worksheet sheet = null;
        private Range lastCell = null;
        private Dictionary<string, Dictionary<string, List<Product>>> categories;
        
        private int row;
        internal Dictionary<string, Dictionary<string, List<Product>>> Categories { get => categories; set => categories = value; }

        
        

        private bool OpenFile(string filename)
        {
            string path = Directory.GetCurrentDirectory() + "\\" + filename;
            excelApp = new Application();
            books = excelApp.Workbooks;
            try
            {
                book = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheets = book.Sheets;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        // Method used for closing excel COM objects
        private void ReleaseFile(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
            }
            catch (Exception)
            {

                throw;
            }
        }
        // Method used for closing excel application
        private void CloseFile()
        {
            // If book was load
            if (book != null)
            {
                ReleaseFile(lastCell);
                foreach (var item in sheets)
                {
                    ReleaseFile(item);
                }
                ReleaseFile(sheets);
                book.Close();
                ReleaseFile(book);
            }
            // Closing application
            books.Close();
            ReleaseFile(books);
            excelApp.Application.Quit();
            excelApp.Quit();
            ReleaseFile(excelApp);
            lastCell = null;
            sheet = null;
            sheets = null;
            book = null;
            books = null;
            excelApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }
        // Metod used for loading data from excel file
        public void SaveFile(string filename)
        {
            OpenFile(filename);
            sheet = (Excel.Worksheet)book.Sheets[1];
            Output output = new Output(sheet);

            object file = excelApp.GetSaveAsFilename("OutputName", "Книга Excel (*.xlsx), *xlsx", Type.Missing, Type.Missing);
            book.SaveAs(file);
            CloseFile();
        }
        public void LoadFile(string filename)
        {
            if (OpenFile(filename))
            {
                LoadCategories();
                CloseFile();
            }
            else
            {
                CloseFile();
                new MessageWindow().ShowMessage("Data file is not exist!");
                Environment.Exit(10);
            }
        }
        private void LoadCategories()
        {
            categories = new Dictionary<string, Dictionary<string, List<Product>>>();
            for (int i = 1; i <= book.Sheets.Count; i++)
            {
                sheet = (Excel.Worksheet)book.Sheets[i];
                Categories.Add(sheet.Name, LoadSubCategories());
            }
        }
        private Dictionary<string, List<Product>> LoadSubCategories()
        {
            // determinate 
            Dictionary<string, List<Product>> subcategories = new Dictionary<string, List<Product>>();
            string subName;

            lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            row = 2;
            while (row <= lastCell.Row)
            {
                
                // enter name of subCategory
                if (sheet.Cells[row, 1].Text.ToString() != "")
                {
                    subName = sheet.Cells[row, 1].Text.ToString();
                    // when sub categoty is end then add sub to dictionary
                    subcategories.Add(subName, LoadProducts());
                }
                
            }
            return subcategories;
        }
        private List<Product> LoadProducts()
        {
            List<Product> productsList = new List<Product>();
            // jump to next row where starts products
            row += 1;
            // repeat until a new subCategory started
            while (sheet.Cells[row, 1].Text.ToString() == "" && row <= lastCell.Row)
            {
                // new instance of Product and get information from table
                Product oneProduct = new Product();
                oneProduct.name = sheet.Cells[row, 2].Text.ToString();
                oneProduct.type = sheet.Cells[row, 3].Text.ToString();
                oneProduct.minPrice = Convert.ToDouble(sheet.Cells[row, 4].Text);
                oneProduct.maxPrice = Convert.ToDouble(sheet.Cells[row, 5].Text);
                // determinate bool variable (if cell empty then false else true)
                if (sheet.Cells[row, 6].Text.ToString() != "")
                    oneProduct.isAct = true;
                // add product to productslist 
                productsList.Add(oneProduct);
                // jump to next row
                row += 1;
            }
            return productsList;
        }

        
    }
}
