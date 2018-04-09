using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf.fonts;
using iTextSharp.text.pdf;
using System.Globalization;
using System.IO;
namespace Human_Resource_Information_System
{
    public partial class PaySlips : Form
    {
        thisDatabase db = new thisDatabase();
        String fileloc_dtr = "";

        private GlobalClass gc;
        private GlobalMethod gm;
        
        
        public PaySlips()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            InitializeComponent();
        }

       
        private void PaySlips_Load(object sender, EventArgs e)
        {
            fileloc_dtr = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            gc.load_employee(cbo_employee);
            gc.load_payroll_period(cbo_payollperiod);
            pic_loading.Visible = false;
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
            pic_loading.Visible = true;
            btn_submit.Enabled = false;
            bgworker.RunWorkerAsync();
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

        private void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            String query = "", empid = "",emp_name = "",employee_id="", date_from = "", date_to = "", pay_code = "", table = "hr_payslip_files", filename = "", code = "", col = "", val = "", date_in = "";
            DataTable pay_period = null;
            

            query = "SELECT empid,firstname,lastname FROM rssys.hr_employee";

            cbo_employee.Invoke(new Action(() =>
            {
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
                date_from = DateTime.Parse(pay_period.Rows[0]["date_from"].ToString()).ToString("yyyy-MM-dd");
                date_to = DateTime.Parse(pay_period.Rows[0]["date_to"].ToString()).ToString("yyyy-MM-dd");
            }
            DateTime StartDate = DateTime.Parse(date_from);
            DateTime EndDate = DateTime.Parse(date_to);
            filename = RandomString(10);
            filename += ".pdf";

            System.IO.FileStream fs = new FileStream(fileloc_dtr + "/ViewController/RPT/Payroll/payslips/" + filename, FileMode.Create);

            

            Document document = new Document(PageSize.A4.Rotate(), 25, 25, 30, 30);

            PdfWriter.GetInstance(document, fs);
            document.Open();

            try
            {
                if (employees.Rows.Count > 0)
                {
                    for(int r = 0; r < employees.Rows.Count; r++)
                    {

                        employee_id = employees.Rows[r]["empid"].ToString();
                        emp_name = employees.Rows[r]["firstname"].ToString() + " " + employees.Rows[r]["lastname"].ToString();
                        PdfPTable pay_table = new PdfPTable(2);
                        float[] widths = new float[] { 30, 70 };
                        pay_table.WidthPercentage = 100f;
                        pay_table.SetWidths(widths);
                        

                        //COLUMN 1 OF PAYSLIP
                        PdfPTable tbt_col_1 = new PdfPTable(2);
                        float[] tbt_col1_row1 = new float[] { 20,20 };
                        tbt_col_1.SetWidths(tbt_col1_row1);
                        tbt_col_1.WidthPercentage = 100f;


                        iTextSharp.text.Font tbt_col1_row1_font = FontFactory.GetFont("Arial", 8);
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("Print Date :", tbt_col1_row1_font)) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph(DateTime.Now.ToString("yyyy-MM-dd"),tbt_col1_row1_font)) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT});

                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 10))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 10))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });



                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("Payslip", FontFactory.GetFont("Arial", 10))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 10))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("Pay Period :", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph(date_from +"-" + date_to, FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("Employee Name :", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph(emp_name, FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });



                        PdfPTable tbt_col1_2nd_row = new PdfPTable(2);
                        tbt_col1_2nd_row.SetWidths(new float[] {10,10});
                        tbt_col1_2nd_row.WidthPercentage = 100f;

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Basic Pay : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Withholding Tax : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Ttl OT Amt : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Ttl Leave Amt : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Std. Deductions : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Ttl Hol. Amt : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Additional Deduct. : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Adjustments : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Additional Deduct. : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });


                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Total Earnings : 0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("Total Deductions :0", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("\n", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("NET PAY : P", FontFactory.GetFont("Arial", 13))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("0.00", FontFactory.GetFont("Arial", 13))) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });

                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("This to acknowledge receipt of my pay for the period " + date_from + " to " + date_to + ". Further acknowledge that in the absence of my written complain within three(3) working days from the date of receipt, amount credited is certified to be final.", FontFactory.GetFont("Arial", 8))) { Colspan = 2, Border = 0, HorizontalAlignment = Element.ALIGN_CENTER,Padding = 10 });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("")) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });


                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("hahahehe", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
                        tbt_col1_2nd_row.AddCell(new PdfPCell(new Paragraph("hahahehedsfasdfas", FontFactory.GetFont("Arial", 8))) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });

                        

                        tbt_col_1.AddCell(new PdfPCell(tbt_col1_2nd_row) { Colspan = 2,Border = 1,HorizontalAlignment = Element.ALIGN_CENTER });
                        tbt_col_1.AddCell(new PdfPCell(new Paragraph("")) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });






                        pay_table.AddCell(new PdfPCell(tbt_col_1));

                        
                        

                        //COLUMN 2 OF PASYLIP

                        PdfPTable tbt_col2_base = new PdfPTable(2);
                        tbt_col2_base.SetWidths(new float[]{50,50 });
                        tbt_col2_base.WidthPercentage = 100f;

                        PdfPTable hdr = new PdfPTable(1);
                        hdr.SetWidths(new float[] {20 });
                        iTextSharp.text.Font hdr_font = FontFactory.GetFont("Arial",13);
                        hdr.AddCell(new PdfPCell(new Paragraph("PAYSLIP",hdr_font)) { Border = 0,HorizontalAlignment = Element.ALIGN_LEFT });
                        tbt_col2_base.AddCell(new PdfPCell(hdr) { Colspan = 2});
                        tbt_col2_base.AddCell(new PdfPCell(new Paragraph("\n")));

                        pay_table.AddCell(new PdfPCell(tbt_col2_base));


                        document.Add(pay_table);
                        document.NewPage();
                    }
                    
                }
            }
            catch
            {
                MessageBox.Show("Operation failed");
            }

            code = db.get_pk("pay_slip_id");
            col = "pay_slip_id,filename,date_created";
            val = "'" + code + "','" + filename + "','" + DateTime.Now.ToShortDateString() + "'";

            if (db.InsertOnTable(table, col, val))
            {
                db.set_pkm99("pay_slip_id", db.get_nextincrementlimitchar(code, 8)); //changes from 'hr_empid'
                MessageBox.Show("New payslips created.");
            }
            else
            {
                MessageBox.Show("Failed on saving.");
            }

            document.Close();
            display_list();
            pic_loading.Invoke(new Action(() => {
                pic_loading.Visible = false;
                btn_submit.Enabled = true;
            }));
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
            dgvl_payslips_files.Invoke(new Action(() =>
            {
                try { dgvl_payslips_files.Rows.Clear(); }
                catch (Exception) { }
                int i = 0;
                String query = "SELECT * FROM rssys.hr_payslip_files ORDER BY date_created";

                try
                {
                    DataTable dt = db.QueryBySQLCode(query);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        i = dgvl_payslips_files.Rows.Add();
                        DataGridViewRow row = dgvl_payslips_files.Rows[i];

                        row.Cells["pay_slip_id"].Value = dt.Rows[r]["pay_slip_id"].ToString();
                        row.Cells["filename"].Value = dt.Rows[r]["filename"].ToString();
                        row.Cells["date_created"].Value = dt.Rows[r]["date_created"].ToString();

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
            String sys_dir = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {
                if (dgvl_payslips_files.Rows.Count > 1)
                {
                    r = dgvl_payslips_files.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_payslips_files["filename", r].Value.ToString();

                        try
                        {
                            System.Diagnostics.Process.Start("AcroRd3d2.exe", sys_dir + "/ViewController/RPT/Payroll/payslips/" + dtr_filename);

                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", sys_dir + "/ViewController/RPT/Payroll/payslips/" + dtr_filename);

                        }
                        catch
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", sys_dir + "/ViewController/RPT/Payroll/payslips/" + dtr_filename);
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

        private void btn_deletefile_Click(object sender, EventArgs e)
        {
            int r = -1;
            String dtr_filename = "", payslip_id = "";
            String sys_dir = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {
                if (dgvl_payslips_files.Rows.Count > 1)
                {
                    r = dgvl_payslips_files.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_payslips_files["filename", r].Value.ToString();
                        payslip_id = dgvl_payslips_files["pay_slip_id", r].Value.ToString();
                        DialogResult result = MessageBox.Show("Are you sure you want to delete this file?", "Confirmation", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            File.Delete(sys_dir + "/ViewController/RPT/Payroll/payslips/" + dtr_filename);
                            String query = "DELETE FROM rssys.hr_payslip_files WHERE pay_slip_id = '" + payslip_id + "'";
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
