using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Human_Resource_Information_System
{
    public class Item_Array
    {
        private String l_item_code = "";
        private String l_item_desc = "";
        private String l_unit = "";
        private String l_recv_qty = "";
        private String l_price = "0.00";

        public String item_code
        {
            get
            {
                return l_item_code;
            }
            set
            {
                l_item_code = value;
            }
        }

        public String item_desc
        {
            get
            {
                return l_item_desc;
            }
            set
            {
                l_item_desc = value;
            }
        }

        public String unit
        {
            get
            {
                return l_unit;
            }
            set
            {
                l_unit = value;
            }
        }

        public String recv_qty
        {
            get
            {
                return l_recv_qty;
            }
            set
            {
                l_recv_qty = value;
            }
        }

        public String price
        {
            get
            {
                return l_price;
            }
            set
            {
                l_price = value;
            }
        }       
    }
}
