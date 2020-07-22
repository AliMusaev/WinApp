using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace winApp.Resources
{
    public class DocumentInfo
    {
        public static string ChekID;
        public static string ChekData;
        public static string ChangeID;
        public static string ChangeData;
        public static Dictionary<string, int> CurrencyNameAndCode = new Dictionary<string, int>() { { "Российский рубль", 643 }, { "Евро", 978 }, { "Доллар США", 840 } };
        public static string VendorName;
        public static string VendorAdress;
        public static string VendorITN;
        public static string VendorRRC;
        public static string ShipperNameAndAdress;
        public static string ConsigneeNameAndAdress;
        public static string DocNumber;
        public static string DocData;
        public static string CustomerName;
        public static string CustomerAdress;
        public static string CustomerITN;
        public static string CustomerRRC;
        public static string CurrencyName;
        public static string CurrencyCode;
        public static string GovermentID;
    }
}
