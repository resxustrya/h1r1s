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
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.fonts;
using iTextSharp.text.pdf.fonts.cmaps;
using System.Globalization;
using System.IO;
namespace Human_Resource_Information_System
{
    public partial class rpt_dtr_summary : Form
    {
        thisDatabase db = new thisDatabase();
        String fileloc_dtr = "";

        private GlobalClass gc;
        private GlobalMethod gm;
       
        public rpt_dtr_summary()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            InitializeComponent();
            
        }

        private void rpt_dtr_summary_Load(object sender, EventArgs e)
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

            bgworker.RunWorkerAsync();
        }


        public Double get_dayoff_ot_total(String empid, String ppid)
        {
            DateTime date_from, date_to;
            String timein = "", timeout = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);
            String result = "";
            String overtime = "";
            String datein = "";
            String query = "";
            Double total = 0;
            String str_dayoff1 = "", str_dayoff_2 = "";
            DataTable dayoff = db.QueryBySQLCode("SELECT dayoff1,dayoff2 FROM rssys.hr_employee WHERE empid ='" + empid + "'");
            DataTable payperiod = db.QueryBySQLCode("SELECT date_from, date_to FROM rssys.hr_payrollpariod WHERE pay_code = '" + ppid + "'");
            DataTable dayoff1 = db.QueryBySQLCode("SELECT dayname FROM rssys.hr_days WHERE day = '" + dayoff.Rows[0]["dayoff1"].ToString() + "'");
            DataTable dayoff2 = db.QueryBySQLCode("SELECT dayname FROM rssys.hr_days WHERE day = '" + dayoff.Rows[0]["dayoff2"].ToString() + "'");
            str_dayoff1 = dayoff1.Rows[0]["dayname"].ToString().ToUpper();
            str_dayoff_2 = dayoff2.Rows[0]["dayname"].ToString().ToUpper();


            if (payperiod.Rows.Count > 0)
            {
                date_from = DateTime.Parse(payperiod.Rows[0]["date_from"].ToString());
                date_to = DateTime.Parse(payperiod.Rows[0]["date_to"].ToString());

                foreach (DateTime day in EachDay(date_from, date_to))
                {
                    if (day.DayOfWeek.ToString().ToUpper() == str_dayoff1 || day.DayOfWeek.ToString().ToUpper() == str_dayoff_2)
                    {
                        query = "SELECT DISTINCT work_date,(SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.work_date = '" + day.ToString("yyyy-MM-dd") + "' AND e.empid ='" + empid + "' ORDER BY work_date ";
                        //System.Diagnostics.Debug.Write(query);
                        DataTable logs = db.QueryBySQLCode(query);
                        if (logs != null && logs.Rows.Count > 0)
                        {
                            for (int r = 0; r < logs.Rows.Count; r++)
                            {
                                if (logs.Rows[r]["timein"].ToString() != "")
                                {
                                    timein = logs.Rows[r]["timein"].ToString();
                                }
                                if (logs.Rows[r]["timeout"].ToString() != "")
                                {
                                    timeout = logs.Rows[r]["timeout"].ToString();
                                }

                                DateTime datetime_out = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timeout);
                                DateTime datetime_in = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timein);
                                int res = DateTime.Compare(datetime_out, datetime_in);

                                if (res > 0)
                                {
                                    TimeSpan diff = datetime_out.Subtract(datetime_in);

                                    total_late = total_late + diff;
                                }
                            }
                        }
                    }
                }
            }
            return Convert.ToDouble(total_late.TotalHours);
        }


        public Double get_legal_hol_ot(String empid, String ppid)
        {

            DateTime date_from, date_to;
            DataTable holiday = null;
            String overtime = "";
            String datein = "";
            Double total = 0.00;
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);
            DataTable legal_hol = db.QueryBySQLCode("SELECT date_from,date_to FROM rssys.hr_payrollpariod WHERE pay_code = '" + ppid + "'");
            if (legal_hol.Rows.Count > 0)
            {
                date_from = DateTime.Parse(legal_hol.Rows[0]["date_from"].ToString());
                date_to = DateTime.Parse(legal_hol.Rows[0]["date_to"].ToString());
                foreach (DateTime day in EachDay(date_from, date_to))
                {
                    try
                    {
                        datein = day.ToString("yyyy-MM-dd");
                        holiday = db.QueryBySQLCode("SELECT date_holiday,holiday_type FROM rssys.hr_holidays WHERE date_holiday = '" + datein + "' AND holiday_type = 'L'");
                        if (holiday.Rows.Count > 0)
                        {
                            overtime = compute_overtime(empid, datein);
                            total += Convert.ToDouble(TimeSpan.Parse(overtime).TotalHours);
                        }
                    }
                    catch
                    {

                    }

                }
            }
            return total;
        }

        private String compute_overtime(String empid, String datein)
        {
            String result = "00:00:00";

            String query = "";
            String timein = "", timeout = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {

                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();


                query = "SELECT DISTINCT e.empid,work_date,(SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.empid = '" + empid + "' AND t.work_date BETWEEN '" + gm.toDateString(datein, "") + "' AND '" + gm.toDateString(datein, "") + "' ORDER BY work_date";

                DataTable logs = db.QueryBySQLCode(query);
                if (logs != null && logs.Rows.Count > 0)
                {
                    for (int r = 0; r < logs.Rows.Count; r++)
                    {
                        timeout = logs.Rows[r]["timeout"].ToString();

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
                }

            }

            return result;
        }
        public Double get_special_hol_ot(String empid, String ppid)
        {
            DateTime date_from, date_to;
            DataTable holiday = null;
            String overtime = "";
            String datein = "";
            Double total = 0;
            DataTable legal_hol = db.QueryBySQLCode("SELECT date_from,date_to FROM rssys.hr_payrollpariod WHERE pay_code = '" + ppid + "'");
            if (legal_hol.Rows.Count > 0)
            {
                date_from = DateTime.Parse(legal_hol.Rows[0]["date_from"].ToString());
                date_to = DateTime.Parse(legal_hol.Rows[0]["date_to"].ToString());
                foreach (DateTime day in EachDay(date_from, date_to))
                {
                    try
                    {
                        datein = day.ToString("yyyy-MM-dd");
                        holiday = db.QueryBySQLCode("SELECT date_holiday,holiday_type FROM rssys.hr_holidays WHERE date_holiday = '" + datein + "' AND holiday_type = 'S' ");
                        if (holiday.Rows.Count > 0)
                        {
                            overtime = compute_overtime(empid, datein);
                            total += Convert.ToDouble(TimeSpan.Parse(overtime).TotalHours);
                        }
                    }
                    catch
                    {

                    }

                }
            }
            return total;
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

        public string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private String compute_daysworked(String empid, String date_from, String date_to)
        {
            String result = "0";
            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            String query = "";

            if (sched.Rows.Count > 0)
            {

                String time_from = sched.Rows[0]["shift_sched_from"].ToString();
                String time_to = sched.Rows[0]["shift_sched_to"].ToString();


                int count = 0;
                query = "SELECT DISTINCT e.empid,work_date,(SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.work_date BETWEEN '" + gm.toDateString(date_from, "") + "' AND '" + gm.toDateString(date_to, "") + "' AND t.empid ='" + empid + "' ORDER BY work_date";

                DataTable logs = db.QueryBySQLCode(query);

                try
                {
                    if (logs != null && logs.Rows.Count > 0)
                    {
                        result = logs.Rows.Count.ToString();
                    }
                    else
                    {
                        result = "0";
                    }
                }
                catch { result = "0"; }

            }
            return result;
        }
        private String compute_absent(String empid, String date_from, String date_to, String total_worked)
        {
            String result = "0";

            int total = 0;
            String query = "SELECT dayoff1,dayoff2 from rssys.hr_employee where empid = '" + empid + "'";
            DataTable dt = db.QueryBySQLCode(query);
            DateTime work_date;
            String dayoff1 = dt.Rows[0]["dayoff1"].ToString();
            String dayoff2 = dt.Rows[0]["dayoff2"].ToString();
            int dayno = 0;
            DateTime StartDate = DateTime.Parse(date_from);
            DateTime EndDate = DateTime.Parse(date_to);

            int off1 = Convert.ToInt32(dayoff1[0].ToString());
            int off2 = Convert.ToInt32(dayoff2[0].ToString());
            //COMPUTER NUMBER OF WORKING DAYS

            foreach (DateTime day in EachDay(StartDate, EndDate))
            {
                dayno = Convert.ToInt32(day.DayOfWeek.ToString("d")) + 1;
                if (dayno != off1 && dayno != off2)
                {
                    total++;
                }
            }


            try
            {
                result = (Convert.ToInt32(total) - Convert.ToInt32(total_worked)).ToString();
                // MessageBox.Show("Total working days :" + total + " Total worked days : " + total_worked + "Result is : " + result);
            }
            catch (Exception ex)
            {
                result = "0";
            }

            return result;
        }

        private Double compute_late(String empid, String date_from, String date_to)
        {
            String result = "00:00:00";

            String query = "";
            String timein = "", timeout = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {
                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();

                query = "SELECT DISTINCT e.empid,work_date,(SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.work_date BETWEEN '" + gm.toDateString(date_from, "") + "' AND '" + gm.toDateString(date_to, "") + "' AND t.empid ='" + empid + "' ORDER BY work_date";

                DataTable logs = db.QueryBySQLCode(query);

                if (logs != null && logs.Rows.Count > 0)
                {
                    for (int r = 0; r < logs.Rows.Count; r++)
                    {
                        timein = logs.Rows[r]["timein"].ToString();

                        DateTime datetime_in = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timein);

                        DateTime datetime_from = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + time_from);

                        int res = DateTime.Compare(datetime_from, datetime_in);

                        if (res < 0)
                        {
                            TimeSpan diff = datetime_in.Subtract(datetime_from);
                            total_late = total_late + diff;
                            result = total_late.ToString();
                        }
                    }
                }

            }
            return Convert.ToDouble(TimeSpan.Parse(result).TotalHours);
        }

        private Double compute_undertime(String empid, String date_from, String date_to)
        {
            String result = "00:00:00";

            String query = "";
            String timein = "", timeout = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {
                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();


                query = "SELECT DISTINCT e.empid,work_date,(SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.empid = '" + empid + "' AND t.work_date BETWEEN '" + gm.toDateString(date_from, "") + "' AND '" + gm.toDateString(date_to, "") + "' ORDER BY work_date";
                //System.Diagnostics.Debug.Write(query);
                DataTable logs = db.QueryBySQLCode(query);
                if (logs != null && logs.Rows.Count > 0)
                {
                    for (int r = 0; r < logs.Rows.Count; r++)
                    {
                        if (logs.Rows[r]["timeout"].ToString() != "")
                        {
                            timeout = logs.Rows[r]["timeout"].ToString();
                        }
                        else
                        {
                            timeout = time_from;
                        }

                        DateTime datetime_out = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timeout);
                        DateTime datetime_to = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + time_to);
                        int res = DateTime.Compare(datetime_to, datetime_out);

                        if (res > 0)
                        {
                            TimeSpan diff = datetime_to.Subtract(datetime_out);
                            //MessageBox.Show("Out Time : " + datetime_to + " Time Out : " + datetime_out + " Late : " + diff);
                            total_late = total_late + diff;
                            result = total_late.ToString();
                        }
                    }
                }
            }

            return Convert.ToDouble(TimeSpan.Parse(result).TotalHours);
        }
        private Double compute_overtime(String empid, String date_from, String date_to)
        {
            String result = "00:00:00";

            String query = "";
            String timein = "", timeout = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if (sched.Rows.Count > 0)
            {

                time_from = sched.Rows[0]["shift_sched_from"].ToString();
                time_to = sched.Rows[0]["shift_sched_to"].ToString();


                query = "SELECT DISTINCT e.empid,work_date,(SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.empid = '" + empid + "' AND t.work_date BETWEEN '" + gm.toDateString(date_from, "") + "' AND '" + gm.toDateString(date_to, "") + "' ORDER BY work_date";

                DataTable logs = db.QueryBySQLCode(query);
                if (logs != null && logs.Rows.Count > 0)
                {
                    for (int r = 0; r < logs.Rows.Count; r++)
                    {
                        timeout = logs.Rows[r]["timeout"].ToString();

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
                }

            }

            return Convert.ToDouble(TimeSpan.Parse(result).TotalHours);
        }
        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }

        private void display_list()
        {
            dgvl_dtr_sum_files.Invoke(new Action(() => {
                try { dgvl_dtr_sum_files.Rows.Clear(); }
                catch (Exception) { }
                int i = 0;
                String query = "SELECT * FROM rssys.hr_dtr_sum_files ORDER BY date_created";

                try
                {
                    DataTable dt = db.QueryBySQLCode(query);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        i = dgvl_dtr_sum_files.Rows.Add();
                        DataGridViewRow row = dgvl_dtr_sum_files.Rows[i];

                        row.Cells["dtr_sum_id"].Value = dt.Rows[r]["dtr_sum_id"].ToString();
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
                if (dgvl_dtr_sum_files.Rows.Count > 1)
                {
                    r = dgvl_dtr_sum_files.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_dtr_sum_files["filename", r].Value.ToString();

                        try
                        {
                            System.Diagnostics.Process.Start("AcroRd3d2.exe", sys_dir + "/ViewController/RPT/TimeKeeping/dtr_summary_pdf/" + dtr_filename);
                            
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", sys_dir + "/ViewController/RPT/TimeKeeping/dtr_summary_pdf/" + dtr_filename);
                            
                        }
                        catch
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", sys_dir + "/ViewController/RPT/TimeKeeping/dtr_summary_pdf/" + dtr_filename);
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
            String dtr_filename = "",dtr_sum_id="";
            String sys_dir = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {
                if (dgvl_dtr_sum_files.Rows.Count > 1)
                {
                    r = dgvl_dtr_sum_files.CurrentRow.Index;

                    try
                    {
                        dtr_filename = dgvl_dtr_sum_files["filename", r].Value.ToString();
                        dtr_sum_id = dgvl_dtr_sum_files["dtr_sum_id", r].Value.ToString();
                        DialogResult result = MessageBox.Show("Are you sure you want to delete this file?", "Confirmation", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            File.Delete(sys_dir + "/ViewController/RPT/TimeKeeping/dtr_summary_pdf/" + dtr_filename);
                            String query = "DELETE FROM rssys.hr_dtr_sum_files WHERE dtr_sum_id = '" + dtr_sum_id +"'";
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

            pic_loading.Invoke(new Action(() =>
            {
                pic_loading.Visible = true;
                btn_submit.Enabled = false;
            }));
            String query = "", empid = "", date_from = "", date_to = "", pay_code = "", table = "hr_dtr_sum_files", filename = "", code = "", col = "", val = "", date_in = "", employee_name = "", days_worked = "", absent_total = "", employee_id = "";

            Double legal_hol_ot = 0.00, special_hol_ot = 0.00, dayoff_ot_total = 0.00, late_total = 0.00, ut_total = 0.00, ot_total = 0.00;

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
            cbo_payollperiod.Invoke(new Action(() =>
            {
                pay_code = cbo_payollperiod.SelectedValue.ToString();
            }));
            

            pay_period = get_date(pay_code);

            if (pay_period.Rows.Count > 0)
            {
                date_from = DateTime.Parse(pay_period.Rows[0]["date_from"].ToString()).ToShortDateString();
                date_to = DateTime.Parse(pay_period.Rows[0]["date_to"].ToString()).ToShortDateString();
            }


            filename = RandomString(5) +"_" + DateTime.Now.ToString("yyyy-MM-dd");
            filename += ".pdf";



            System.IO.FileStream fs = new FileStream(fileloc_dtr + "/ViewController/RPT/TimeKeeping/dtr_summary_pdf/" + filename, FileMode.Create);



            Document document = new Document(PageSize.LEGAL.Rotate(), 25, 25, 30, 30);

            PdfWriter.GetInstance(document, fs);
            document.Open();

            BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 9, iTextSharp.text.Font.NORMAL);

            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = Element.ALIGN_CENTER;
            paragraph.Font = FontFactory.GetFont("Arial", 12);
            paragraph.SetLeading(1, 1);
            paragraph.Add("DAILY TIME RECORD SUMMARY");



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
            float[] widths = new float[] { 10, 10, 10, 10, 10, 10, 10, 10, 10};
            t.WidthPercentage = 100;
            t.SetWidths(widths);

            t.AddCell(new PdfPCell(new Phrase(new Chunk("Employee Name", font))) { HorizontalAlignment = Element.ALIGN_LEFT });
            t.AddCell(new PdfPCell(new Phrase(new Chunk("Days Worked (# days)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
            t.AddCell(new PdfPCell(new Phrase(new Chunk("Abcences (# days)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
            t.AddCell(new PdfPCell(new Phrase(new Chunk("Total Late (Hrs/Amt.)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });

            t.AddCell(new PdfPCell(new Phrase(new Chunk("Total Undertime(Hrs/Amt.)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //t.AddCell(new PdfPCell(new Phrase(new Chunk("Total Overtime", font))) { HorizontalAlignment = Element.ALIGN_CENTER });

           t.AddCell(new PdfPCell(new Phrase(new Chunk("Reg OT(Hrs/Amt.)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
            t.AddCell(new PdfPCell(new Phrase(new Chunk("Dayoff OT(Hrs/Amt.)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
            t.AddCell(new PdfPCell(new Phrase(new Chunk("Legal Hol. OT(Hrs/Amt.)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
            t.AddCell(new PdfPCell(new Phrase(new Chunk("Special Hol. OT(Hrs/Amt.)", font))) { HorizontalAlignment = Element.ALIGN_CENTER });



            for (int r = 0; r < employees.Rows.Count; r++)
            {
                legal_hol_ot = 0.00;
                special_hol_ot = 0.00;
                employee_id = employees.Rows[r]["empid"].ToString();
                employee_name = employees.Rows[r]["firstname"].ToString() + ", " + employees.Rows[r]["lastname"].ToString();
                days_worked = compute_daysworked(employee_id, date_from, date_to);
                absent_total = compute_absent(employee_id, date_from, date_to, days_worked);
                late_total = compute_late(employee_id, date_from, date_to);
                ut_total = compute_undertime(employee_id, date_from, date_to);
                ot_total = compute_overtime(employee_id, date_from, date_to);
                dayoff_ot_total = get_dayoff_ot_total(employee_id, pay_code);


                legal_hol_ot = get_legal_hol_ot(employee_id, pay_code);
                special_hol_ot = get_special_hol_ot(employee_id, pay_code);
              



                t.AddCell(new PdfPCell(new Phrase(new Chunk(employee_name, font))) { HorizontalAlignment = Element.ALIGN_LEFT });
                t.AddCell(new PdfPCell(new Phrase(new Chunk(days_worked, font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                t.AddCell(new PdfPCell(new Phrase(new Chunk(absent_total, font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                if (late_total > 0)
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk(late_total.ToString("0.00"), font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                else
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk("", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                if (ut_total > 0)
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk(ut_total.ToString("0.00"), font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                else
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk("", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                
                if (ot_total > 0)
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk(ot_total.ToString("0.00"), font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                else
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk("", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                
                //t.AddCell(new PdfPCell(new Phrase(new Chunk("0", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                if (dayoff_ot_total > 0)
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk(dayoff_ot_total.ToString("0.00"), font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                else
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk("", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }

                if (legal_hol_ot > 0)
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk(legal_hol_ot.ToString("0.00"), font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                else
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk("", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                if (special_hol_ot > 0)
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk(special_hol_ot.ToString("0.00"), font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }
                else
                {
                    t.AddCell(new PdfPCell(new Phrase(new Chunk("", font))) { HorizontalAlignment = Element.ALIGN_CENTER });
                }

            }


            document.Add(t);
            document.Close();


            code = db.get_pk("dtr_sum_id");
            col = "dtr_sum_id,filename,date_created";
            val = "'" + code + "','" + filename + "','" + DateTime.Now.ToShortDateString() + "'";

            if (db.InsertOnTable(table, col, val))
            {
                db.set_pkm99("dtr_sum_id", db.get_nextincrementlimitchar(code, 8)); //changes from 'hr_empid'
                MessageBox.Show("New DTR Summary created.");
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

      
    }
}
