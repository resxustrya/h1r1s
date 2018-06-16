using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Human_Resource_Information_System
{
    public partial class rpt_payroll_summary : Form
    {
        thisDatabase db = new thisDatabase();
        String fileloc_dtr = "";

        private GlobalClass gc;
        private GlobalMethod gm;
        public rpt_payroll_summary()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            InitializeComponent();
        }

        private void rpt_payroll_summary_Load(object sender, EventArgs e)
        {
            fileloc_dtr = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            pic_loading.Visible = false;
            gc.load_payroll_period(cbo_payollperiod);
            gc.load_dept(cbo_department);
            gc.load_employee(cbo_employee);
            display_list();
        }

        public string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void display_list()
        {
            dgvl_payroll_summary.Invoke(new Action(() => {
                try { dgvl_payroll_summary.Rows.Clear(); }
                catch (Exception) { }
                int i = 0;
                String query = "SELECT * FROM rssys.hr_rpt_files WHERE rpt_type = 'PAYROLL_SUMMARY' ORDER BY date_added";

                try
                {
                    DataTable dt = db.QueryBySQLCode(query);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        i = dgvl_payroll_summary.Rows.Add();
                        DataGridViewRow row = dgvl_payroll_summary.Rows[i];

                        row.Cells["rpt_id"].Value = dt.Rows[r]["rpt_id"].ToString();
                        row.Cells["filename"].Value = dt.Rows[r]["filename"].ToString();
                        row.Cells["date_added"].Value = dt.Rows[r]["date_added"].ToString();

                        i++;
                    }
                }
                catch { }
            }));
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

        private void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            String query = "",dept_query = "",deptid = "", empid = "", date_from = "", date_to = "", pay_code = "", table = "hr_rpt_files", filename = "", code = "", col = "", val = "", date_in = "";
            DataTable pay_period = null;
            Double total = 0.00, ee = 0.00, er = 0.00, ec = 0.00;
            
            query = "SELECT empid, firstname, lastname,mi,pay_rate,department FROM rssys.hr_employee";
            dept_query = "SELECT deptid FROM rssys.hr_department";

            cbo_employee.Invoke(new Action(() => {
                if (cbo_employee.SelectedIndex != -1)
                {
                    empid = cbo_employee.SelectedValue.ToString();
                    query += " WHERE empid='" + empid + "'";
                }
            }));

            cbo_department.Invoke(new Action(() =>
            {
                if(cbo_department.SelectedIndex != -1)
                {
                    deptid = cbo_department.SelectedValue.ToString();
                    dept_query += " WHERE deptid = '" + deptid + "' LIMIT 1";
                }
            }));
            query += " ORDER BY empid ASC";

            DataTable departments = db.QueryBySQLCode(dept_query);
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


            //System.IO.FileStream fs = new FileStream("\\\\RIGHTAPPS\\RightApps\\Eastland\\payroll_reports\\payroll_summary\\" + filename, FileMode.Create);
            System.IO.FileStream fs = new FileStream(fileloc_dtr + "\\ViewController\\RPT\\Payroll\\payroll_summary\\" + filename, FileMode.Create);

            Document document = new Document(PageSize.LEGAL.Rotate());

            PdfWriter.GetInstance(document, fs);
            document.Open();

            document.Add(new Paragraph("SUGBU REALTY, INC", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_LEFT });

            document.Add(new Paragraph("PAYROLL PERIOD", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_LEFT });

            document.Add(new Chunk("\n"));

            PdfPTable _thead = new PdfPTable(13);
            _thead.WidthPercentage = 100f;
            float[] columnWidths = { 100, 40,40, 40, 40, 40, 40, 40, 40, 40, 40,40,40 };
            _thead.SetWidths(columnWidths);



            _thead.AddCell(new PdfPCell(new Paragraph("Employee", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Rate", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Basic Pay", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Late/UT", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Overtime", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Gross Pay", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Earnings", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Deductions", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("SSS", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("PHIC", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("HDMF", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("With-Tax-C", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
            _thead.AddCell(new PdfPCell(new Paragraph("Net Pay", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });

            cbo_employee.Invoke(new Action(() =>
            {
                if (cbo_employee.SelectedIndex != -1)
                {
                    foreach (DataRow _emp in employees.Rows)
                    {
                        DataTable emp_dept = db.QueryBySQLCode("SELECT * FROM rssys.hr_department WHERE deptid='" + _emp["department"].ToString() + "' LIMIT 1");
                        DataTable emp_payroll = db.QueryBySQLCode("SELECT * FROM rssys.hr_emp_payroll WHERE empid = '" + _emp["empid"].ToString() + "'");

                        _thead.AddCell(new PdfPCell(new Paragraph(emp_dept.Rows[0]["dept_name"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLDITALIC)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_LEFT });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });


                        Double pay_rate = Convert.ToDouble(_emp["pay_rate"].ToString());
                        Double basic_pay = Convert.ToDouble(emp_payroll.Rows[0]["basic_pay"].ToString());
                        Double late_ut = Convert.ToDouble(emp_payroll.Rows[0]["late_ut"].ToString());



                        _thead.AddCell(new PdfPCell(new Paragraph(_emp["lastname"].ToString() + ", " + _emp["firstname"].ToString() + " " + _emp["mi"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_LEFT });
                        _thead.AddCell(new PdfPCell(new Paragraph(pay_rate.ToString("N", new CultureInfo("en-US")), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_RIGHT });
                        _thead.AddCell(new PdfPCell(new Paragraph(basic_pay.ToString("N", new CultureInfo("en-US")), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_RIGHT });
                        _thead.AddCell(new PdfPCell(new Paragraph("-" +late_ut.ToString("N", new CultureInfo("en-US")), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });
                        _thead.AddCell(new PdfPCell(new Paragraph("", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD)) { Alignment = Element.ALIGN_CENTER }) { HorizontalAlignment = Element.ALIGN_CENTER });


                        //DataTable _emp_payroll = db.QueryBySQLCode("SELECT ");

                    }
                }
            }));


            document.Add(_thead);


            document.Close();
            code = db.get_pk("rpt_id");
            col = "rpt_id,filename,date_added,rpt_type";
            val = "'" + code + "','" + filename + "','" + DateTime.Now.ToShortDateString() + "','PAYROLL_SUMMARY'";

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

        private void btn_print_Click(object sender, EventArgs e)
        {
            int r = -1;
            String dtr_filename = "";
            //String sys_dir = "\\\\RIGHTAPPS\\RightApps\\Eastland\\payroll_reports\\payroll_summary\\";
            String sys_dir = fileloc_dtr + "\\ViewController\\RPT\\Payroll\\payroll_summary\\";
            try
            {
                if (dgvl_payroll_summary.Rows.Count > 1)
                {
                    r = dgvl_payroll_summary.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_payroll_summary["filename", r].Value.ToString();

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
                    MessageBox.Show("Files is empty.");
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void btn_deletefile_Click(object sender, EventArgs e)
        {
            int r = -1;
            String dtr_filename = "", rpt_id = "";
            //String sys_dir = "\\\\RIGHTAPPS\\RightApps\\Eastland\\payroll_reports\\payroll_summary\\";
            String sys_dir = fileloc_dtr + "\\ViewController\\RPT\\Payroll\\payroll_summary\\";

            try
            {
                if (dgvl_payroll_summary.Rows.Count > 1)
                {
                    r = dgvl_payroll_summary.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_payroll_summary["filename", r].Value.ToString();
                        rpt_id = dgvl_payroll_summary["rpt_id", r].Value.ToString();
                        DialogResult result = MessageBox.Show("Are you sure you want to delete this file?", "Confirmation", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            File.Delete(sys_dir + dtr_filename);
                            String query = "DELETE FROM rssys.hr_rpt_files WHERE rpt_id = '" + rpt_id + "' AND rpt_type = 'PAYROLL_SUMMARY'";
                            db.QueryBySQLCode(query);
                            MessageBox.Show("File successfully deleted");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to remove file. It may not exist");
                    }
                }
                else
                {
                    MessageBox.Show("Empty files.");
                }
            }
            catch (Exception ex)
            {

            }
            display_list();
        }
    }
}
