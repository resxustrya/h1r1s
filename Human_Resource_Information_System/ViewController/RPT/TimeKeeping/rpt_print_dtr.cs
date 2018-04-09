using System;
using System.IO;
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
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.fonts;
using iTextSharp.text.pdf.fonts.cmaps;
using System.Globalization;

namespace Human_Resource_Information_System
{
    public partial class rpt_print_dtr : Form
    {
        thisDatabase db = new thisDatabase();
        String fileloc_dtr = "";
        
        private GlobalClass gc;
        private GlobalMethod gm;
        public rpt_print_dtr()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            
            InitializeComponent();
        }

        private void rpt_print_dtr_Load(object sender, EventArgs e)
        {
            pic_loading.Visible = false;
            fileloc_dtr = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            
            gc.load_employee(cbo_employee);
            gc.load_payroll_period(cbo_payollperiod);
            display_list();
        }
        private void display_list()
        {
            dgvl_dtrfiles.Invoke(new Action(() => {
                try { dgvl_dtrfiles.Rows.Clear(); }
                catch (Exception) { }
                int i = 0;
                String query = "SELECT * FROM rssys.hr_dtr_files ORDER BY date_created";

                try
                {
                    DataTable dt = db.QueryBySQLCode(query);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        i = dgvl_dtrfiles.Rows.Add();
                        DataGridViewRow row = dgvl_dtrfiles.Rows[i];

                        row.Cells["dtr_id"].Value = dt.Rows[r]["dtr_id"].ToString();
                        row.Cells["filename"].Value = dt.Rows[r]["filename"].ToString();
                        row.Cells["date_created"].Value = dt.Rows[r]["date_created"].ToString();

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

        private String compute_undertime(String empid, String timeout,String datein)
        {
            String result = "";

            String query = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {
                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();


                DateTime datetime_out = Convert.ToDateTime(datein + " " + timeout);
                DateTime datetime_to = Convert.ToDateTime(datein + " " + time_to);
                int res = DateTime.Compare(datetime_to, datetime_out);

                if (res > 0)
                {
                    TimeSpan diff = datetime_to.Subtract(datetime_out);
                    //MessageBox.Show("Out Time : " + datetime_to + " Time Out : " + datetime_out + " Late : " + diff);
                    total_late = total_late + diff;
                    result = total_late.ToString();
                }
                  
            }

            return result;
        }
        private String compute_late(String empid,String timein,String datein)
        {
            String result = "";

           
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {
                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();

                

                DateTime datetime_in = Convert.ToDateTime(datein + " " + timein);

                DateTime datetime_from = Convert.ToDateTime(datein + " " + time_from);

                int res = DateTime.Compare(datetime_from, datetime_in);

                if (res < 0)
                {
                    TimeSpan diff = datetime_in.Subtract(datetime_from);
                    total_late = total_late + diff;
                    result = total_late.ToString();
                }
            }
            return result;
        }
        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }


        private String compute_overtime(String empid, String timeout, String datein)
        {
            String result = "00:00:00";

            String query = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {

                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();

               
                DateTime datetime_out = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timeout);
                DateTime datetime_to = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + time_to);
                int res = DateTime.Compare(datetime_to, datetime_out);

                if (res < 0)
                {
                    TimeSpan diff = datetime_out.Subtract(datetime_to);
                    //   MessageBox.Show("Out Time : " + datetime_to + " Time Out : " + datetime_out + " Overtime : " + diff);

                    total_late = total_late + diff;
                    result = total_late.ToString();
                }
            }

            return result;
        }

        public string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
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

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter_1(object sender, EventArgs e)
        {

        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            int r = -1;
            String dtr_filename = "";
            String sys_dir = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            
            try
            {
                if (dgvl_dtrfiles.Rows.Count > 1)
                {
                    r = dgvl_dtrfiles.CurrentRow.Index;
                    try
                    {
                        dtr_filename = dgvl_dtrfiles["filename", r].Value.ToString();
                        
                        try
                        {
                            System.Diagnostics.Process.Start("chrome.exe", sys_dir + "/ViewController/RPT/TimeKeeping/dtr_pdf/" + dtr_filename);
                        }
                        catch(Exception ex)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", sys_dir + "/ViewController/RPT/TimeKeeping/dtr_pdf/" + dtr_filename);
                        }
                        catch
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", sys_dir + "/ViewController/RPT/TimeKeeping/dtr_pdf/" + dtr_filename);
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
            String dtr_filename = "", dtr_sum_id = "";
            String sys_dir = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {
                if (dgvl_dtrfiles.Rows.Count > 1)
                {
                    r = dgvl_dtrfiles.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_dtrfiles["filename", r].Value.ToString();
                        dtr_sum_id = dgvl_dtrfiles["dtr_id", r].Value.ToString();
                        DialogResult result = MessageBox.Show("Are you sure you want to delete this file?", "Confirmation", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            File.Delete(sys_dir + "/ViewController/RPT/TimeKeeping/dtr_pdf/" + dtr_filename);
                            String query = "DELETE FROM rssys.hr_dtr_files WHERE dtr_id = '" + dtr_sum_id + "'";
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
        private void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            String query = "", empid = "", date_from = "", date_to = "", pay_code = "", table = "hr_dtr_files", filename = "", code = "", col = "", val = "", date_in = "";
            DataTable pay_period = null;
            

            query = "SELECT empid,firstname,lastname FROM rssys.hr_employee";
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
                date_from = DateTime.Parse(pay_period.Rows[0]["date_from"].ToString()).ToShortDateString();
                date_to = DateTime.Parse(pay_period.Rows[0]["date_to"].ToString()).ToShortDateString();
            }
            DateTime StartDate = DateTime.Parse(date_from);
            DateTime EndDate = DateTime.Parse(date_to);
            try
            {


                filename = RandomString(5) + "_" + DateTime.Now.ToString("yyyy-MM-dd");
                filename += ".pdf";

                System.IO.FileStream fs = new FileStream(fileloc_dtr + "/ViewController/RPT/TimeKeeping/dtr_pdf/" + filename, FileMode.Create);
                Document document = new Document(PageSize.LEGAL, 25, 25, 30, 30);

                PdfWriter.GetInstance(document, fs);
                document.Open();
                if (employees.Rows.Count > 0)
                {

                    for (int r = 0; r < employees.Rows.Count; r++)
                    {

                        Paragraph paragraph = new Paragraph();
                        paragraph.Alignment = Element.ALIGN_CENTER;
                        paragraph.Font = FontFactory.GetFont("Arial", 12);
                        paragraph.SetLeading(1, 1);
                        paragraph.Add("DAILY TIME RECORD");
                        Phrase line_break = new Phrase("\n");
                        document.Add(paragraph);
                        document.Add(line_break);

                        empid = employees.Rows[r]["empid"].ToString();




                        Paragraph emp_name = new Paragraph();
                        emp_name.Alignment = Element.ALIGN_CENTER;
                        emp_name.Font = FontFactory.GetFont("Arial", 8);
                        emp_name.SetLeading(1, 1);
                        emp_name.Add(employees.Rows[r]["firstname"].ToString() + " " + employees.Rows[r]["lastname"].ToString());
                        document.Add(emp_name);

                        Paragraph horizontal_line = new Paragraph();
                        horizontal_line.Alignment = Element.ALIGN_CENTER;
                        horizontal_line.Font = FontFactory.GetFont("Arial", 10);
                        horizontal_line.SetLeading(1, 1);
                        horizontal_line.Add("--------------------------------------------------------------------------------------");
                        document.Add(horizontal_line);

                        Paragraph label_name = new Paragraph();
                        label_name.Alignment = Element.ALIGN_CENTER;
                        label_name.SetLeading(1, 1);
                        label_name.Font = FontFactory.GetFont("Arial", 8);
                        label_name.Add("Name");

                        document.Add(label_name);

                        Phrase line_break_1 = new Phrase("\n");
                        line_break_1.SetLeading(0.5f, 0.5f);
                        document.Add(line_break_1);

                        PdfPTable t = new PdfPTable(9);
                        float[] widths = new float[] { 10, 20, 20, 20, 20, 20, 20, 20, 20 };
                        t.WidthPercentage = 100;
                        t.SetWidths(widths);
                        t.AddCell(new PdfPCell(new Phrase("DATE")) { Colspan = 2, Rowspan = 2, HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("AM")) { Colspan = 2, HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("PM")) { Colspan = 2, HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("UT/OT")) { Colspan = 3, HorizontalAlignment = Element.ALIGN_CENTER });


                        t.AddCell(new PdfPCell(new Phrase("IN")) { HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("OUT")) { HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("IN")) { HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("OUT")) { HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("LATE")) { HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("UT")) { HorizontalAlignment = Element.ALIGN_CENTER });
                        t.AddCell(new PdfPCell(new Phrase("OT")) { HorizontalAlignment = Element.ALIGN_CENTER });



                        query = "SELECT DISTINCT CONCAT(lastname,' ',firstname) AS name, t.source, e.empid, to_char(work_date, 'yyyy-MM-dd') AS work_date, (SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE e.empid = '" + empid + "' AND t.work_date BETWEEN '" + date_from + "' AND '" + date_to + "' ORDER BY work_date ";


                        DataTable dt = db.QueryBySQLCode(query);

                        String date_name = "", am_in = "", am_out = "", pm_in = "", pm_out = "", log_date = "", late = "", ut = "", ot_total = "";
                        int index = 0;
                        foreach (DateTime day in EachDay(StartDate, EndDate))
                        {
                            if (index != dt.Rows.Count)
                            {
                                if (day.ToShortDateString() == DateTime.Parse(dt.Rows[index]["work_date"].ToString()).ToShortDateString())
                                {
                                    log_date = DateTime.Parse(dt.Rows[index]["work_date"].ToString()).ToShortDateString();
                                    am_in = dt.Rows[index]["timein"].ToString();
                                    pm_out = dt.Rows[index]["timeout"].ToString();
                                    index++;
                                }
                            }
                            date_name = day.ToString("MMM d, yyyy ddd", CultureInfo.InvariantCulture);

                            if (day.ToShortDateString() != log_date)
                            {
                                am_in = "";
                                pm_out = "";
                            }

                            t.AddCell(new PdfPCell(new Phrase(date_name)) { Colspan = 2, HorizontalAlignment = Element.ALIGN_LEFT });
                            t.AddCell(new PdfPCell(new Phrase(am_in)) { HorizontalAlignment = Element.ALIGN_CENTER });
                            t.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_CENTER });
                            t.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_CENTER });
                            t.AddCell(new PdfPCell(new Phrase(pm_out)) { HorizontalAlignment = Element.ALIGN_CENTER });
                            if (am_in != "")
                            {
                                late = compute_late(empid, am_in, day.ToShortDateString());
                            }
                            else { late = ""; }


                            t.AddCell(new PdfPCell(new Phrase(late)) { HorizontalAlignment = Element.ALIGN_CENTER });
                            if (pm_out != "")
                            {
                                ut = compute_undertime(empid, pm_out, day.ToShortDateString());
                            }
                            else { ut = ""; }



                            t.AddCell(new PdfPCell(new Phrase(ut)) { HorizontalAlignment = Element.ALIGN_CENTER });

                            if (pm_out != "")
                            {
                                ot_total = compute_overtime(empid, pm_out, day.ToShortDateString());
                            }

                            t.AddCell(new PdfPCell(new Phrase(ot_total)) { HorizontalAlignment = Element.ALIGN_CENTER });
                            ut = "";
                            late = "";
                            am_in = "";
                            pm_out = "";
                            ot_total = "";
                        }

                        document.Add(t);
                        document.NewPage();
                    }
                }


                document.Close();
                code = db.get_pk("dtr_id");
                col = "dtr_id,filename,date_created";
                val = "'" + code + "','" + filename + "','" + DateTime.Now.ToShortDateString() + "'";

                if (db.InsertOnTable(table, col, val))
                {
                    db.set_pkm99("dtr_id", db.get_nextincrementlimitchar(code, 8)); //changes from 'hr_empid'
                    MessageBox.Show("DTR PRINTED");
                }
                else
                {
                    MessageBox.Show("Failed on saving.");
                }
            }
            catch
            {
                MessageBox.Show("Operation failed");
            }
            //bgworker.RunWorkerAsync();
            pic_loading.Invoke(new Action(() => {
                pic_loading.Visible = false;
            }));

            btn_submit.Invoke(new Action(() => {
                btn_submit.Enabled = true;
            }));
            display_list();
        }
    }
}
