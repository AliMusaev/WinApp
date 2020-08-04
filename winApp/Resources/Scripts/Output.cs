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
        Excel.Worksheet sheet;
        public Output(Excel.Worksheet sheet)
        {
            this.sheet = sheet;
        }
        public void LoadCalculatedData(Dictionary<int, double> result, List<Product> products, double cost)
        {
            // first row number 
            int k = 25;
            foreach (var item in result)
            {
                if(item.Key * item.Value < cost)
                {
                    if(item.Key > 0)
                    {
                        InsertRow(k, sheet);
                        sheet.Cells[k, "A"] = products[k - 25].name;
                        char[] arr = products[k - 25].name.ToCharArray();
                        int y = 1;
                        if (arr.Length > 18)
                            y = arr.Length / 18;
                        sheet.Rows[k].RowHeight = 24 * y;
                        sheet.Cells[k, "AI"] = products[k - 25].type;
                        sheet.Cells[k, "DH"] = item.Key * item.Value;
                        sheet.Cells[k, "BB"] = item.Value;
                        sheet.Cells[k, "AT"] = item.Key;
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
            cellRange = null;
            cellRange1 = null;
            rowRange = null;
            //cellRange = sheet.get_Range("A" + (rowNum+1), "FF" + (rowNum+1));
            //cellRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

        }
    }
}
