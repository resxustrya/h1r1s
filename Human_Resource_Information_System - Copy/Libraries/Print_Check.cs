using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing;
using System.Data;

namespace Human_Resource_Information_System
{
    class Print_Check
    {
        PrintDocument pdoc = null;
        DateTime t_date;
        String payto = "";
        Double amt = 0;
        

        public Print_Check()
        {

        }

        public Print_Check(String payto, Double amt, DateTime t_date)
        {
            this.payto = payto;
            this.amt = amt;
            this.t_date = t_date;

            print();
        }

        public void print()
        {
            PrintDialog pd = new PrintDialog();
            pdoc = new PrintDocument();
            PrinterSettings ps = new PrinterSettings();
            //Font font = new Font("Arial", 12);


            /*PaperSize psize = new PaperSize("Custom", 850, 350);
            ps.DefaultPageSettings.PaperSize = psize;

            pd.Document = pdoc;
            pd.Document.DefaultPageSettings.PaperSize = psize;
            pdoc.DefaultPageSettings.PaperSize.Height = 850;
            pdoc.DefaultPageSettings.PaperSize.Height = 350;

            */
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
            GlobalMethod gm = new GlobalMethod();
            NumberToEnglish_orig amtinwords = new NumberToEnglish_orig();
            GlobalClass gc = new GlobalClass();
            Graphics graphics = e.Graphics;
            String fontfamily = "Arial";
            Font font = new Font(fontfamily, 10);
            //Font font_headernote = new Font(fontfamily, 9);
            //Font font_bd = new Font(fontfamily, 8, FontStyle.Italic);
           // Font font_grand = new Font(fontfamily, 9, FontStyle.Underline);
            int ppwidth = pdoc.DefaultPageSettings.PaperSize.Width;
            float fontHeight = font.GetHeight();
            int startX = 0;
            int startY = 0;
            int Offset = 15;
            int incY = 30;
            int col2 = 30;
            int col3 = 120;
            int col4 = 625;
            String str = "";

            Offset = Offset + incY;
            str = t_date.ToString("MMM dd, yyyy");
            graphics.DrawString(str, font, new SolidBrush(Color.Black), startX + col4, Offset);

            Offset = Offset + incY;
            graphics.DrawString(this.payto, font, new SolidBrush(Color.Black), startX + col3, Offset);

            graphics.DrawString(gm.toAccountingFormat(this.amt), font, new SolidBrush(Color.Black), startX + col4, Offset);

            Offset = Offset + incY;

            graphics.DrawString(amtinwords.changeCurrencyToWordsForCheck(this.amt.ToString()), font, new SolidBrush(Color.Black), startX + col3 - 10, Offset);
        }
    }
}