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
                if (lastCell != null)
                    ReleaseFile(lastCell);
                foreach (var item in sheets)
                {
                    ReleaseFile(item);
                }
                ReleaseFile(sheets);
                book.Close(false);
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

        public void SaveFile(List<Result> result, List<Product> products, double cost)
        {
            OpenFile("Form.xlsx");
            sheet = book.Sheets[1];
            LoadCalculatedData(result, products, cost);
            object file = excelApp.GetSaveAsFilename("OutputName", "Книга Excel (*.xlsx), *xlsx", Type.Missing, Type.Missing);
            book.SaveCopyAs(file);
            CloseFile();
        }
        void LoadCalculatedData(List<Result> result, List<Product> products, double cost)
        {
            EnterFelds();
            // first row number 
            int k = 25;
            foreach (var item in result)
            {
                if (item.Amount * item.Price <= cost)
                {
                    if (item.Amount > 0)
                    {
                        InsertRow(k);
                        sheet.Cells[k, "A"] = products[k - 25].name;
                        char[] arr = products[k - 25].name.ToCharArray();
                        int y = 1;
                        if (arr.Length > 18)
                            y = arr.Length / 18;
                        sheet.Rows[k].RowHeight = 24 * y;
                        sheet.Cells[k, "AI"] = products[k - 25].type;
                        sheet.Cells[k, "DH"] = item.Amount * item.Price;
                        sheet.Cells[k, "BB"] = item.Price;
                        sheet.Cells[k, "AT"] = item.Amount;
                        k++;
                    }
                }
            }
            sheet.Cells[k + 1, "DH"] = cost;
        }
        void EnterFelds()
        {
            sheet.Cells[7, "AF"] = DocumentInfo.ChekID;
            sheet.Cells[7, "BF"] = DocumentInfo.ChekData;
            sheet.Cells[8, "AF"] = DocumentInfo.ChangeID;
            sheet.Cells[8, "BF"] = DocumentInfo.ChangeData;
            sheet.Cells[10, "M"] = DocumentInfo.VendorName;
            sheet.Cells[11, "I"] = DocumentInfo.VendorAdress;
            sheet.Cells[12, "Y"] = DocumentInfo.VendorITN + "/" + DocumentInfo.VendorRRC;
            sheet.Cells[13, "AI"] = DocumentInfo.ShipperNameAndAdress;
            sheet.Cells[14, "AH"] = DocumentInfo.ConsigneeNameAndAdress;
            sheet.Cells[15, "AR"] = DocumentInfo.DocNumber;
            sheet.Cells[15, "BO"] = DocumentInfo.DocData;
            sheet.Cells[16, "O"] = DocumentInfo.CustomerName;
            sheet.Cells[17, "I"] = DocumentInfo.CustomerAdress;
            sheet.Cells[18, "AA"] = DocumentInfo.CustomerITN + "/" + DocumentInfo.CustomerRRC;
            sheet.Cells[19, "AG"] = DocumentInfo.CurrencyName + ", " + DocumentInfo.CurrencyCode;
            sheet.Cells[20, "CP"] = DocumentInfo.GovermentID;
        }
        void InsertRow(int rowNum)
        {
            //Excel.Range cellRange = sheet.get_Range("A"+rowNum,"FF"+rowNum);
            //cellRange.Copy(Type.Missing);
            Excel.Range cellRange1 = sheet.Cells[rowNum, 1];
            Excel.Range rowRange = cellRange1.EntireRow;
            rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, true);
            Excel.Range cellRange = sheet.get_Range("A" + rowNum, "T" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("U" + rowNum, "AB" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("AC" + rowNum, "AH" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("AI" + rowNum, "AS" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("AT" + rowNum, "BA" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("BB" + rowNum, "BL" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("BM" + rowNum, "CA" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("CB" + rowNum, "CK" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("CL" + rowNum, "CU" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("CV" + rowNum, "DG" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("DH" + rowNum, "DV" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("DW" + rowNum, "EF" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("EG" + rowNum, "ES" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = sheet.get_Range("ET" + rowNum, "FE" + rowNum);
            cellRange.Merge(Type.Missing);
            cellRange = null;
            cellRange1 = null;
            rowRange = null;
            //cellRange = sheet.get_Range("A" + (rowNum+1), "FF" + (rowNum+1));
            //cellRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

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
                sheet = book.Sheets[i];
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
