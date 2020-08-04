using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace winApp.Resources
{
    class SubCategories
    {
        Dictionary<string, List<Product>> data;
        public Dictionary<string, List<Product>> Data { get => data; private set => data = value; }

        public Dictionary<string, List<Product>> LoadSubCategories(Excel.Workbook book , string sheetName)
        {
            Data = new Dictionary<string, List<Product>>();
            string subName = null;
            Excel.Worksheet sheet = null;

            // Find desired sheet 
            for (int i = 1; i <= book.Sheets.Count; i++)
            {
                if (book.Sheets[i].Name == sheetName)
                {
                    sheet = (Excel.Worksheet)book.Sheets[i];
                    break;
                }
            }
            // Determinate last cells horizontal and vertical
            var lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            // Start counting rows ignoring column names
            int j = 2;
            // Repeat to the end of the table
            while (j <= lastCell.Row)
            {

                // determinate 
                List<Product> subCategoryList = new List<Product>();
                // enter name of subCategory
                if (sheet.Cells[j, 1].Text.ToString() != "")
                {
                    // Get subName from table
                    subName = sheet.Cells[j, 1].Text.ToString();
                    // jump to next row where starts products
                    j += 1;
                    // repeat until a new subCategory started
                    while (sheet.Cells[j, 1].Text.ToString() == "" && j <= lastCell.Row)
                    {
                        // new instance of Product and get information from table
                        Product oneProduct = new Product();
                        oneProduct.name = sheet.Cells[j, 2].Text.ToString();
                        oneProduct.type = sheet.Cells[j, 3].Text.ToString();
                        oneProduct.minPrice = Convert.ToDouble(sheet.Cells[j, 4].Text);
                        oneProduct.maxPrice = Convert.ToDouble(sheet.Cells[j, 5].Text);
                        // determinate bool variable (if cell empty then false else true)
                        if (sheet.Cells[j, 6].Text.ToString() != "")
                            oneProduct.isAct = true;
                        // add product to productslist 
                        subCategoryList.Add(oneProduct);
                        // jump to next row
                        j += 1;
                    }
                    // when sub categoty is end then add sub to dictionary
                    Data.Add(subName, subCategoryList);
                }
            }
            return Data;
        }
    }

}
