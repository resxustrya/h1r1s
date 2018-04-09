using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Human_Resource_Information_System
{
    class GenericClass
    {

    }

    public class enterFromPOtoReceiving
    {
        public String line { get; set; }
        public string itemCode { get; set; }
        public string description { get; set; }
        public string qty { get; set; }
        public string unit { get; set; }
        public string cost { get; set; } //price
        public string lineAmnt { get; set; }
        public string discount { get; set; }
        public string ln_vat { get; set; }
        public string cc_code { get; set; }
        public string scc_code { get; set; }
        public string po_num { get; set; }
        public string po_line { get; set; }
        public string non_vat { get; set; }
    }

    public class enterFromPRtoPO
    {
        public String line { get; set; }
        public string itemCode { get; set; }
        public string description { get; set; }
        public string qty { get; set; }
        public string unit { get; set; }
        public string cost { get; set; }
        public string lineAmnt { get; set; }
        public string cc_code { get; set; }
        public string scc_code { get; set; }
        public string pr_num { get; set; }
        public string pr_reference { get; set; }

    }
    public class enterRetPurItems
    {
        public String llno {get; set;}
        public String itemcode { get; set; }
        public String itemdesc { get; set; }
        public String purchUnit { get; set; }
        public String purchUnitSelectedValue { get; set; }
        public String qty { get; set; }
        public String costprice { get; set; }
        public String lnamt { get; set; }
    }
}
