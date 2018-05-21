using System;
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
    public partial class rpt_loan_summary : Form
    {
        thisDatabase db = new thisDatabase();
        String fileloc_dtr = "";
        private GlobalClass gc;
        private GlobalMethod gm;
        public rpt_loan_summary()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            InitializeComponent();
        }

        private void rpt_loan_summary_Load(object sender, EventArgs e)
        {
            fileloc_dtr = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            pic_loading.Visible = false;
            gc.load_payroll_period(cbo_payollperiod);
            gc.load_employee(cbo_employee);
            gc.load_loan_type(cbo_loan_type);
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

        private void btn_print_Click(object sender, EventArgs e)
        {

        }

        private void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            String query = "", empid = "",loan_type ="", date_from = "", date_to = "", pay_code = "", table = "hr_rpt_files", filename = "", code = "", col = "", val = "", date_in = "";
            DataTable pay_period = null;

            Double total = 0.00;
            String loan_query = "";
            query = "SELECT empid, firstname, lastname,pay_rate FROM rssys.hr_employee";
            cbo_employee.Invoke(new Action(() => {
                if (cbo_employee.SelectedIndex != -1)
                {
                    empid = cbo_employee.SelectedValue.ToString();
                    query += " WHERE empid='" + empid + "'";
                }
            }));
            query += " ORDER BY empid ASC";


            DataTable employees = db.QueryBySQLCode(query);

            loan_query = "SELECT * FROM hr_loan_type";
            cbo_loan_type.Invoke(new Action(() =>
            {
                if(cbo_loan_type.SelectedIndex != -1)
                {
                    loan_type = cbo_loan_type.SelectedValue.ToString();
                    loan_query += " WHERE code = '" + loan_type + "'";
                }
            }));
            DataTable dt_loan_type = db.QueryBySQLCode(loan_query);

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
            System.IO.FileStream fs = new FileStream(fileloc_dtr + "\\ViewController\\RPT\\Payroll\\loan_summary\\" + filename, FileMode.Create);

            Document document = new Document(PageSize.LEGAL, 25, 25, 30, 30);

            PdfWriter.GetInstance(document, fs);
            document.Open();

            BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 9, iTextSharp.text.Font.NORMAL);

            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = Element.ALIGN_CENTER;
            paragraph.Font = FontFactory.GetFont("Arial", 12);
            paragraph.SetLeading(1, 1);
            paragraph.Add("EMPLOYEE LOAN SUMMARY REPORTS");


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



            PdfPTable t = new PdfPTable(1);
            float[] widths = new float[] { 100 };
            t.WidthPercentage = 100;
            t.SetWidths(widths);

            PdfPTable dis_earnings = new PdfPTable(4);
            float[] _w2 = new float[] { 70, 50, 50, 50 };
            dis_earnings.SetWidths(_w2);
            dis_earnings.AddCell(new PdfPCell(new Paragraph("Employee Name")) { Border = 2 });
            dis_earnings.AddCell(new PdfPCell(new Paragraph("EmployeE Share")) { Border = 2 });
            dis_earnings.AddCell(new PdfPCell(new Paragraph("EmployeR Share")) { Border = 2 });
            dis_earnings.AddCell(new PdfPCell(new Paragraph("Total Contribution")) { Border = 2 });

            foreach (DataRow _employees in employees.Rows)
            {
                String fname = _employees["firstname"].ToString();
                String lname = _employees["lastname"].ToString();
                String empno = _employees["empid"].ToString();
                try
                {
                    DataTable has_earnings = db.QueryBySQLCode("SELECT emp_no FROM rssys.hr_loanhdr WHERE employee_no = '" + empno + "'");
                    if (has_earnings != null && has_earnings.Rows.Count > 0)
                    {
                        dis_earnings.AddCell(new PdfPCell(new Paragraph(fname + " " + lname)) { Colspan = 2, Border = 2 });

                        DataTable hoe = db.QueryBySQLCode("SELECT DISTINCT(de.code) as code, de.description  FROM rssys.hr_loan_type de  LEFT JOIN rssys.hr_loanhdr od  ON de.code = od.loan_type WHERE od.employee_no = '" + empno + "'");

                        foreach (DataRow _hoe in hoe.Rows)
                        {
                            dis_earnings.AddCell(new PdfPCell(new Paragraph(_hoe["description"].ToString())) { PaddingLeft = 30f, Colspan = 2, Border = 0 });
                            DataTable hee = db.QueryBySQLCode("SELECT * FROM rssys.hr_loanhdr WHERE loan_type = '" + _hoe["code"].ToString() +"' AND emp_no = '" + empno + "'");
                            foreach (DataRow _hee in hee.Rows)
                            {
                                dis_earnings.AddCell(new PdfPCell(new Paragraph(_hee["loan_amount"].ToString())) { PaddingLeft = 40f, Colspan = 2, Border = 0 });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            document.Add(dis_earnings);
            document.Add(t);
            document.Close();

            code = db.get_pk("rpt_id");
            col = "rpt_id,filename,date_added,rpt_type";
            val = "'" + code + "','" + filename + "','" + DateTime.Now.ToShortDateString() + "','LOAN'";

            if (db.InsertOnTable(table, col, val))
            {
                db.set_pkm99("rpt_id", db.get_nextincrementlimitchar(code, 8)); //changes from 'hr_empid'
                MessageBox.Show("New summary reports created");
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
            dgvl_loan_summary.Invoke(new Action(() => {
                try { dgvl_loan_summary.Rows.Clear(); }
                catch (Exception) { }
                int i = 0;
                String query = "SELECT * FROM rssys.hr_rpt_files WHERE rpt_type = 'LOAN' ORDER BY date_added";

                try
                {
                    DataTable dt = db.QueryBySQLCode(query);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        i = dgvl_loan_summary.Rows.Add();
                        DataGridViewRow row = dgvl_loan_summary.Rows[i];

                        row.Cells["rpt_id"].Value = dt.Rows[r]["rpt_id"].ToString();
                        row.Cells["filename"].Value = dt.Rows[r]["filename"].ToString();
                        row.Cells["date_added"].Value = dt.Rows[r]["date_added"].ToString();

                        i++;
                    }
                }
                catch { }
            }));
        }
        public string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
