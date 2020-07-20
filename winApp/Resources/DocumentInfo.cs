using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace winApp.Resources
{
    public class DocumentInfo
    {
        public Dictionary<string, int> CurrencyNameAndCode = new Dictionary<string, int>() { { "Российский рубль", 643 }, { "Евро", 978 }, { "Доллар США", 840 } };
        public string VendorName;
        public string VendorAdress;
        public int VendorITN;
        public int VendorRRC;
        public string ShipperNameAndAdress;
        public string ConsigneeNameAndAdress;
        public string DocNumber;
        public string DocData;
        public string CustomerName;
        public string CustomerAdress;
        public int CustomerITN;
        public int CustomerRRC;
        public string CurrencyName;
        public int CurrencyCode;
        public string GovermentID;
    }
}
