using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing;
using System.Data;
/*
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;*/
using System.Drawing.Printing;
using System.Text;
using System.Drawing;
using System.IO;
using System.ComponentModel;

namespace Human_Resource_Information_System
{
    class BarcodePrintOut
    {
        PrintDocument pdoc = null;
        String _item1 = null, _item2 = null, _item3 = null;
        int _noofcopies = 1;
        thisDatabase db;
        int _c1 = 0, _c2 = 0, _c3 = 0;

        public BarcodePrintOut(String item1, String item2, String item3, int copies, int c1, int c2, int c3)
        {
            db = new thisDatabase();

            _item1 = item1;
            _item2 = item2;
            _item3 = item3;
            _noofcopies = copies;
            _c1 = c1;
            _c2 = c2;
            _c3 = c3;
        }
        public void print()
        {
            try
            {
                PrintDocument pd = new PrintDocument();
                pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);

                Margins margins = new Margins(0, 0, 0, 0);
                pd.DefaultPageSettings.Margins = margins;
               
                // Set the printer name. 
                //pd.PrinterSettings.PrinterName = "\\NS5\hpoffice
                //pd.PrinterSettings.PrinterName = "Zebra New GK420t"               
                pd.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString());
            }

            MessageBox.Show("Done Printing.");
        }

        void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            Font printFont = new Font("IDAutomationHC39M", 6, FontStyle.Regular);
            //Font printFont = new Font("Code128", 6, FontStyle.Regular);
            Font printFont1 = new Font("Arial Narrow", 7, FontStyle.Bold);

            SolidBrush br = new SolidBrush(Color.Black);
            int y = 5;
            String ccode1 = db.get_item_costcode(_item1), desc1 = db.get_item_desc(_item1), price1 = db.get_itemregsellprice(_item1).ToString("0.00");
            String ccode2 = db.get_item_costcode(_item2), desc2 = db.get_item_desc(_item2), price2 = db.get_itemregsellprice(_item2).ToString("0.00");
            String ccode3 = db.get_item_costcode(_item3), desc3 = db.get_item_desc(_item3), price3 = db.get_itemregsellprice(_item3).ToString("0.00");

            int c1 = _c1, c2 = _c2, c3 = _c3;
            float endPointX1 = 0, endPointX2 = 0, endPointX3 = 0, endPointY = 0;
            // price
            //cost code

            for (int i = 1; i <= _noofcopies; i++)
            {
                // item description
                endPointY = y + 35;
                endPointX1 = 130;
                endPointX2 =  130;
                endPointX3 = 130;

                ev.Graphics.DrawString("*" + desc1 + "*", printFont1, br, new RectangleF(_c1, y, endPointX1, endPointY));
                ev.Graphics.DrawString("*" + desc2 + "*", printFont1, br, new RectangleF(_c2, y, endPointX2, endPointY));
                ev.Graphics.DrawString("*" + desc3 + "*", printFont1, br, new RectangleF(_c3, y, endPointX3, y+35));
                y = y + 35;

                // barcode
                ev.Graphics.DrawString("*" + _item1 + "*", printFont, br, _c1, y);
                ev.Graphics.DrawString("*" + _item2 + "*", printFont, br, _c2, y);
                ev.Graphics.DrawString("*" + _item3 + "*", printFont, br, _c3, y);
                y = y + 30;

                // barcode
                ev.Graphics.DrawString("*" + _item1 + "*", printFont1, br, _c1, y);
                ev.Graphics.DrawString("*" + _item2 + "*", printFont1, br, _c2, y);
                ev.Graphics.DrawString("*" + _item3 + "*", printFont1, br, _c3, y);
                y = y + 10;

                //price
                ev.Graphics.DrawString("*PHP " + price1 + "*", printFont1, br, _c1, y);
                ev.Graphics.DrawString("*PHP " + price2 + "*", printFont1, br, _c2, y);
                ev.Graphics.DrawString("*PHP " + price3 + "*", printFont1, br, _c3, y);

                //ccode
                ev.Graphics.DrawString("*" + ccode1 + "*", printFont1, br, _c1 + 40, y);
                ev.Graphics.DrawString("*" + ccode2 + "*", printFont1, br, _c2 + 40, y);
                ev.Graphics.DrawString("*" + ccode3 + "*", printFont1, br, _c3 + 40, y);                

                y = y + 15;
                

                //ev.Graphics.DrawString("*" + _item1 + "*", printFont1, br, _c1, y);
                //ev.Graphics.DrawString("*" + _item2 + "*", printFont1, br, _c2, y);
                //ev.Graphics.DrawString("*" + _item3 + "*", printFont1, br, _c3, y);

                //y = y + 15;
            }
                
        }
    }
}
