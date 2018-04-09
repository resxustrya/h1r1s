using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing;
using System.Data;

namespace Human_Resource_Information_System{
    class Receipt
    {
        PrintDocument pdoc = null;
        String checkNo;
        DateTime t_date;
        String cust_name, dtype, clerk, disc_name, tbl_no, rmno, mcoupon, t_time, pax, senior;
        Double grand_amnt, subtotal, disc_amnt, meal, sc, vat_sc, vat_nonsc;
        DataTable dt;
        Boolean has_disc = false, has_reg;
        Boolean has_sc;

        public Boolean _has_disc
        {
            set { this.has_disc = value; }
            get { return this.has_disc; }
        }

        public Boolean _has_sc
        {
            set { this.has_sc = value; }
            get { return this.has_sc; }
        }

        public String _checkNo
        {
            set { this.checkNo = value; }
            get { return this.checkNo; }
        }

        public String _clerk
        {
            set { this.clerk = value; }
            get { return this.clerk; }
        }

        public DateTime _t_date
        {
            set { this.t_date = value; }
            get { return this.t_date; }
        }

        public String _t_time
        {
            set { this.t_time = value; }
            get { return this.t_time; }
        }

        public String _tbl_no
        {
            set { this.tbl_no = value; }
            get { return this.tbl_no; }
        }

        public String _rmno
        {
            set { this.rmno = value; }
            get { return this.rmno; }
        }

        public String _mcoupon
        {
            set { this.mcoupon = value; }
            get { return this.mcoupon; }
        }

        public String _cust_name
        {
            set { this.cust_name = value; }
            get { return this.cust_name; }
        }

        public String _pax
        {
            set { this.pax = value; }
            get { return this.pax; }
        }

        public String _senior
        {
            set { this.senior = value; }
            get { return this.senior; }
        }

        public String _dtype
        {
            set { this.dtype = value; }
            get { return this.dtype; }
        }

        public DataTable _dt
        {
            set { this.dt = value; }
            get { return this.dt; }
        }

        public String _disc_name
        {
            set { this.disc_name = value; }
            get { return this.disc_name; }
        }

        public Double _disc_amnt
        {
            set { this.disc_amnt = value; }
            get { return this.disc_amnt; }
        }

        public Double _meal
        {
            set { this.meal = value; }
            get { return this.meal; }
        }

        public Double _sc
        {
            set { this.sc = value; }
            get { return this.sc; }
        }

        public Double _vat_sc
        {
            set { this.vat_sc = value; }
            get { return this.vat_sc; }
        }

        public Double _vat_nonsc
        {
            set { this.vat_nonsc = value; }
            get { return this.vat_nonsc; }
        }

        public Double _subtotal
        {
            set { this.subtotal = value; }
            get { return this.subtotal; }
        }

        public Double _grand_amnt
        {
            set { this.grand_amnt = value; }
            get { return this.grand_amnt; }
        }

        public Receipt()
        {
            checkNo = "XXXXXXXX";
            t_date = DateTime.Now;
            cust_name = "CUSTOMER NAME"; dtype = "DINE IN"; clerk = "COMPUTER";
            t_time = DateTime.Now.ToString("hh:mm tt");
            grand_amnt = 0.00;
            dt = null;
            has_disc = false; has_reg = false;
            has_sc = false;
        }

        public Receipt(String checkNo, DateTime lt_date, String cust_name, String dtype, float grand_amnt, String clerk)
        {
            this.checkNo = checkNo;
            this.t_date = lt_date;
            this.cust_name = cust_name;
            this.dtype = dtype;
            this.grand_amnt = grand_amnt;
            this.clerk = clerk;
        }

        public void print()
        {
            PrintDialog pd = new PrintDialog();
            pdoc = new PrintDocument();
            PrinterSettings ps = new PrinterSettings();
            Font font = new Font("Courier New", 15);


            //PaperSize psize = new PaperSize("Custom", 100, 200);
            //ps.DefaultPageSettings.PaperSize = psize;

            //pd.Document = pdoc;
            //pd.Document.DefaultPageSettings.PaperSize = psize;
            //pdoc.DefaultPageSettings.PaperSize.Height =320;
            //pdoc.DefaultPageSettings.PaperSize.Height = 820;


            //pdoc.DefaultPageSettings.PaperSize.Width = 255;

            pdoc.PrintPage += new PrintPageEventHandler(pdoc_PrintPage);

            DialogResult result = pd.ShowDialog();
            //if (result == DialogResult.OK)
            //{
            PrintPreviewDialog pp = new PrintPreviewDialog();
            pp.Document = pdoc;
            pp.Show();
            //DialogResult result = pp.ShowDialog();
            //if (result == DialogResult.OK)
            //{

            //pdoc.Print();
            //}
            //}
        }

        void pdoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics graphics = e.Graphics;
            String fontfamily = "Fake Receipt";
            Font font = new Font(fontfamily, 8);
            Font font_headernote = new Font(fontfamily, 9);
            Font font_bd = new Font(fontfamily, 8, FontStyle.Italic);
            Font font_grand = new Font(fontfamily, 9, FontStyle.Underline);
            int ppwidth = pdoc.DefaultPageSettings.PaperSize.Width;
            float fontHeight = font.GetHeight();
            int startX = 0;
            int startY = 0;
            int Offset = 10;
            int incY = 15;
            int col2 = 30;
            int col3 = 150;
            int col4 = 200;
            int bdX = 10;
            String str = "";
            int desc_length = 0;
            String underLine = "----------------------------------------------------------";
            String cust_name = this.cust_name;
            String clerk = this.clerk;

            Offset = Offset + incY;
            str = "SO# " + checkNo;
            graphics.DrawString(str, font, new SolidBrush(Color.Black), startX, startY + Offset);

            graphics.DrawString(t_date.ToString("MMM dd yyyy (ddd)"), font, new SolidBrush(Color.Black), startX + col3, startY + Offset);
            Offset = Offset + incY;

            graphics.DrawString("Customer: " + cust_name, font, new SolidBrush(Color.Black), startX, startY + Offset);
            Offset = Offset + incY;


            //header column
            Offset = Offset + incY;
            graphics.DrawString("QTY", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString("DESCRIPTION", font, new SolidBrush(Color.Black), startX + col2, startY + Offset);
            graphics.DrawString("PRICE", font, new SolidBrush(Color.Black), startX + col3, startY + Offset);
            graphics.DrawString("TOTAL", font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString(underLine, font, new SolidBrush(Color.Black), startX, startY + Offset);
            Offset = Offset + incY;
            //end of header column

            //items content
            if (dt != null)
            {
                foreach (DataRow row in dt.Rows)
                {
                    desc_length = row[1].ToString().Length - 1;

                    if (desc_length > 18)
                    {
                        desc_length = 18;
                    }

                    graphics.DrawString(row[0].ToString(), font, new SolidBrush(Color.Black), startX, startY + Offset);
                    graphics.DrawString(row[1].ToString().Substring(0, desc_length), font, new SolidBrush(Color.Black), startX + col2, startY + Offset);
                    graphics.DrawString(row[2].ToString(), font, new SolidBrush(Color.Black), startX + col3, startY + Offset);
                    graphics.DrawString(row[3].ToString(), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
                    Offset = Offset + incY;
                }
            }
            //end of item content

            Offset = Offset + incY;
            Offset = Offset + incY;
            Offset = Offset + incY;
            Offset = Offset + incY;
            graphics.DrawString("TOTAL", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(subtotal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("CASH", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(subtotal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("CHANGE", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(subtotal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;

            graphics.DrawString(underLine, font, new SolidBrush(Color.Black), startX, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("VAT Sale", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(":" + meal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("VAT-Exempt Sale", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(":" + meal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("Total Sale", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(":" + meal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("12% VAT", font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(":" + meal.ToString("0.00"), font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            Offset = Offset + incY;
            graphics.DrawString(clerk, font, new SolidBrush(Color.Black), startX, startY + Offset);
            graphics.DrawString(t_time, font, new SolidBrush(Color.Black), startX + col4, startY + Offset);
            Offset = Offset + incY;
            graphics.DrawString("0 item(s)", font, new SolidBrush(Color.Black), startX + col4, startY + Offset + 5);
            Offset = Offset + incY;
            Offset = Offset + incY;
            graphics.DrawString("Thank you for coming. Please come again.", font, new SolidBrush(Color.Black), startX, startY + Offset);
            Offset = Offset + incY;
        }
    }
}
