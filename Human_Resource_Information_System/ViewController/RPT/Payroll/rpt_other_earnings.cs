﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.fonts;
using iTextSharp.text.pdf.fonts.cmaps;
using System.Globalization;
using System.IO;

using System.Windows.Forms;

namespace Human_Resource_Information_System
{
    public partial class rpt_other_earnings : Form
    {
        thisDatabase db = new thisDatabase();
        String fileloc_dtr = "";

        private GlobalClass gc;
        private GlobalMethod gm;
        public rpt_other_earnings()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            InitializeComponent();
        }

        private void OtherEarnings_Load(object sender, EventArgs e)
        {
            fileloc_dtr = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            pic_loading.Visible = false;
            gc.load_payroll_period(cbo_payollperiod);
            gc.load_employee(cbo_employee);
            display_list();
        }

        private void btn_submit_Click(object sender, EventArgs e)
        {
            if (cbo_payollperiod.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a payroll period.");
                cbo_payollperiod.DroppedDown = true;
                return;
            }
            btn_submit.Enabled = false;
            pic_loading.Visible = true;
            bgworker.RunWorkerAsync();
        }

        public string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            String query = "", empid = "", date_from = "", date_to = "", pay_code = "", table = "hr_other_earnings_files", filename = "", code = "", col = "", val = "", date_in = "";
            DataTable pay_period = null;


            query = "SELECT empid, firstname, lastname FROM rssys.hr_employee";
            cbo_employee.Invoke(new Action(() => {
                if (cbo_employee.SelectedIndex != -1)
                {
                    empid = cbo_employee.SelectedValue.ToString();
                    query += " WHERE empid='" + empid + "'";
                }
            }));

            query += " ORDER BY empid ASC";

            DataTable employees = db.QueryBySQLCode(query);
            cbo_payollperiod.Invoke(new Action(() => {
                pay_code = cbo_payollperiod.SelectedValue.ToString();
            }));

            pay_period = get_date(pay_code);

            if (pay_period.Rows.Count > 0)
            {
                date_from = gm.toDateString(pay_period.Rows[0]["date_from"].ToString(), "yyyy-MM-dd");
                date_to = gm.toDateString(pay_period.Rows[0]["date_to"].ToString(), "yyyy-MM-dd");
            }

            filename = RandomString(5) + "_" + DateTime.Now.ToString("yyyy-MM-dd");
            filename += ".pdf";


            //System.IO.FileStream fs = new FileStream("\\\\RIGHTAPPS\\RightApps\\Eastland\\payroll_reports\\other_earnings\\" + filename, FileMode.Create);
            System.IO.FileStream fs = new FileStream(fileloc_dtr + "\\ViewController\\RPT\\Payroll\\other_earnings\\" + filename, FileMode.Create);



            Document document = new Document(PageSize.LEGAL, 25, 25, 30, 30);

            PdfWriter.GetInstance(document, fs);
            document.Open();

            BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 9, iTextSharp.text.Font.NORMAL);

            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = Element.ALIGN_CENTER;
            paragraph.Font = FontFactory.GetFont("Arial", 12);
            paragraph.SetLeading(1, 1);
            paragraph.Add("OTHER EARNINGS/INCOME SUMMARY");



            Phrase line_break = new Phrase("\n");
            document.Add(paragraph);
            document.Add(line_break);

            Paragraph paragraph_2 = new Paragraph();
            paragraph_2.Alignment = Element.ALIGN_CENTER;
            paragraph_2.Font = FontFactory.GetFont("Arial", 12);
            paragraph_2.SetLeading(1, 1);
            paragraph_2.Add("For the payroll period " + date_from + " to " + date_to);


            Phrase line_break_2 = new Phrase("\n");
            document.Add(paragraph_2);
            document.Add(line_break_2);


            PdfPTable t = new PdfPTable(9);
            float[] widths = new float[] { 100};
            t.WidthPercentage = 100;
            t.SetWidths(widths);

            

            for (int r = 0; r < employees.Rows.Count; r++)
            {
                
            }
            

            document.Add(t);
            document.Close();
            code = db.get_pk("earning_id");
            col = "earning_id,filename,date_added";
            val = "'" + code + "','" + filename + "','" + DateTime.Now.ToShortDateString() + "'";

            if (db.InsertOnTable(table, col, val))
            {
                db.set_pkm99("earning_id", db.get_nextincrementlimitchar(code, 8)); //changes from 'hr_empid'
                MessageBox.Show("New Other Earnings Summary added.");
            }
            else
            {
                MessageBox.Show("Failed on saving.");
            }
            pic_loading.Invoke(new Action(() => {
                pic_loading.Visible = false;
                btn_submit.Enabled = true;
            }));

            display_list();

        }

        private DataTable get_date(String code)
        {
            DataTable dt = null;
            try
            {
                dt = db.QueryBySQLCode("SELECT date_from,date_to from rssys.hr_payrollpariod where pay_code='" + code + "'");
            }
            catch { }
            return dt;
        }

        private void display_list()
        {
            dgvl_other_earnings.Invoke(new Action(() => {
                try { dgvl_other_earnings.Rows.Clear(); }
                catch (Exception) { }
                int i = 0;
                String query = "SELECT * FROM rssys.hr_other_earnings_files ORDER BY date_added";

                try
                {
                    DataTable dt = db.QueryBySQLCode(query);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        i = dgvl_other_earnings.Rows.Add();
                        DataGridViewRow row = dgvl_other_earnings.Rows[i];

                        row.Cells["earning_id"].Value = dt.Rows[r]["earning_id"].ToString();
                        row.Cells["filename"].Value = dt.Rows[r]["filename"].ToString();
                        row.Cells["date_added"].Value = dt.Rows[r]["date_added"].ToString();

                        i++;
                    }
                }
                catch { }
            }));

        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            int r = -1;
            String dtr_filename = "";
            //String sys_dir = "\\\\RIGHTAPPS\\RightApps\\Eastland\\payroll_reports\\dtr_summary\\";
            String sys_dir = fileloc_dtr + "\\ViewController\\RPT\\Payroll\\other_earnings\\";
            try
            {
                if (dgvl_other_earnings.Rows.Count > 1)
                {
                    r = dgvl_other_earnings.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_other_earnings["filename", r].Value.ToString();

                        try
                        {
                            System.Diagnostics.Process.Start("AcroRd3d2.exe", sys_dir + dtr_filename);

                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", sys_dir + dtr_filename);
                        }
                        catch
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", sys_dir + dtr_filename);
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Select a filename");
                    }
                }
                else
                {
                    MessageBox.Show("DTR files is empty.");
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}