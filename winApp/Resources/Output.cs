using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace winApp.Resources
{
    class Output
    {
        List <string> sheets;
        Microsoft.Office.Interop.Excel.Application excelApp;
        Excel.Workbook Book;
        Excel.Worksheet sheet;
        public Output()
        {
            sheets = new List<string>();
            string path = Directory.GetCurrentDirectory() + "\\Form.xlsx";
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            Book = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing);
            for (int j = 1; j <= Book.Sheets.Count; j++)
            {
                sheets.Add(Book.Sheets[j].Name);
            }

            sheet = (Excel.Worksheet)Book.Sheets[1];
            var lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            EnterFelds();
        }
        public void LoadCalculatedData(List<Product> products, double cost)
        {
            int k = 25;
            InsertRow(k, sheet);
            sheet.Cells[k, "A"] = products[0].name;
            char[] arr = products[0].name.ToCharArray();
            int y = arr.Length / 18;
            sheet.Rows[25].RowHeight = 24 * y;
            sheet.Cells[k, "AI"] = products[0].type;
            sheet.Cells[k, "DH"] = cost;
            sheet.Cells[k, "BB"] = cost;
            sheet.Cells[k, "AT"] = 1;
            k++;
            sheet.Cells[k + 1, "DH"] = cost;
        }
        public void LoadCalculatedData(List<List<int>> results, List<Product> products, List<double> calculatedPrices, double cost)
        {
            Random rand = new Random();
            int z = rand.Next(0, results.Count - 1);
            // first row number 
            int k = 25;
            for (int i = 0; i < products.Count; i++)
            {
                if (results[z][i] * calculatedPrices[i] < cost)
                {
                    if (results[z][i] > 0)
                    {

                        InsertRow(k, sheet);
                        sheet.Cells[k, "A"] = products[i].name;
                        char[] arr = products[i].name.ToCharArray();
                        int y = arr.Length / 18;
                        sheet.Rows[k].RowHeight = 24 * y;
                        sheet.Cells[k, "AI"] = products[i].type;
                        sheet.Cells[k, "DH"] = (results[z][i]) * calculatedPrices[i];
                        sheet.Cells[k, "BB"] = calculatedPrices[i];
                        sheet.Cells[k, "AT"] = results[z][i];
                        k++;

                    }
                    sheet.Cells[k + 1, "DH"] = cost;
                }
            }

            
        }
        public void OutputExit()
        {
            Book.SaveCopyAs(Directory.GetCurrentDirectory() + "\\Output.xlsx");
            Book.Close(false);
            excelApp.Quit();
            System.Windows.MessageBox.Show("Complete", "Done", MessageBoxButton.OK);
        }
        void EnterFelds()
        {
            sheet.Cells[7, "AF"] = DocumentInfo.ChekID;
            sheet.Cells[7, "BF"] = DocumentInfo.ChekData;
            sheet.Cells[8, "AF"] = DocumentInfo.ChangeID;
            sheet.Cells[8, "BF"] = DocumentInfo.ChangeData;
            sheet.Cells[10,"M"] = DocumentInfo.VendorName;
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
        void InsertRow(int rowNum, Excel.Worksheet sheet)
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


            //cellRange = sheet.get_Range("A" + (rowNum+1), "FF" + (rowNum+1));
            //cellRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

        }
    }
}
