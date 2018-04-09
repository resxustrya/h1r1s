using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Human_Resource_Information_System
{
    public partial class p_GeneratePayroll : Form
    {

      
        private GlobalClass gc;
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
       
        public p_GeneratePayroll()
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            InitializeComponent();
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            bg_worker.RunWorkerAsync();
        }

        private void bg_worker_DoWork(object sender, DoWorkEventArgs e)
        {
            String table = "hr_emp_payroll", col = "", val = "", emp_pay_code = "";
            String summ_code = "", empid = "", days_worked = "", absences = "", late = "", undertime = "", overtime = "", ppid = "";
            Double total_late = 0.00, total_under_time = 0.00, total_overtime = 0.00, legal_hol_ot = 0.00, special_hol_ot=0.00,dayoff_ot_total = 0.00;
            DataTable dtr = get_generated_dtr();
            Boolean success = false;
            int bar = 1;

            DataTable holidays = null;
            if (dtr.Rows.Count > 0)
            {
                //try
                //{
                    pbar.Invoke(new Action(() =>
                    {
                        pbar.Maximum = dtr.Rows.Count;
                    }));
                    for (int r = 0; r < dtr.Rows.Count; r++)
                    {
                        legal_hol_ot = 0;
                        special_hol_ot = 0;
                        ppid = dtr.Rows[r]["ppid"].ToString();
                        empid = dtr.Rows[r]["empid"].ToString();
                        days_worked = dtr.Rows[r]["days_worked"].ToString();
                        absences = dtr.Rows[r]["absences"].ToString();
                        late = dtr.Rows[r]["late"].ToString();
                        undertime = dtr.Rows[r]["undertime"].ToString();
                        overtime = dtr.Rows[r]["total_overtime"].ToString();

                        legal_hol_ot = get_legal_hol_ot(empid,ppid);
                        special_hol_ot = get_special_hol_ot(empid, ppid);
                        dayoff_ot_total = get_dayoff_ot_total(empid, ppid);
                        try
                        {
                            total_late = Convert.ToDouble(TimeSpan.Parse(late).TotalHours);
                            Double total_late_min = Convert.ToDouble(TimeSpan.Parse(late).TotalHours);
                        }
                        catch (Exception ex)
                        {
                            total_late = 0.00;
                        }

                        try
                        {
                            total_under_time = Convert.ToDouble(TimeSpan.Parse(undertime).TotalHours);
                        }
                        catch (Exception ex)
                        {
                            total_under_time = 0.00;
                        }
                        try
                        {
                            total_overtime = Convert.ToDouble(TimeSpan.Parse(overtime).TotalHours);
                        }
                        catch (Exception ex)
                        {
                            total_overtime = 0.00;
                        }

                        emp_pay_code = db.get_pk("emp_pay_code");
                        col = "emp_pay_code,empid,days_worked,abcences,late,undertime,overtime,regular_ot_a,ppid,legal_hol_ot_a,special_hol_ot_a,dayoff_ot_a";
                        val = "'" + emp_pay_code + "','" + empid + "','" + days_worked + "','" + absences + "','" + total_late.ToString("0.00") + "','" + total_under_time.ToString("0.00") + "','" + total_overtime.ToString("0.00") + "','" + total_overtime.ToString("0.00") + "','" + ppid + "','" +legal_hol_ot.ToString("0.00") + "','" + special_hol_ot.ToString("0.00") +"','" + dayoff_ot_total.ToString("0.00") + "'";

                        if (db.InsertOnTable(table, col, val))
                        {
                            col = "isgenerated='1'";

                            if (db.UpdateOnTable("hr_dtr_sum_employees", col, "empid='" + empid + "' AND ppid='" + ppid + "'"))
                            {
                                success = true;
                                inc_pbar(bar, dtr.Rows.Count);
                                bar++;
                            }
                            db.set_pkm99("emp_pay_code", db.get_nextincrementlimitchar(emp_pay_code, 8));
                        }
                    }
                    if (success)
                    {
                        String period = get_payrol_period(ppid);
                        MessageBox.Show("New payroll was generated From " + period + " .");
                    }
               // }
               // catch(Exception ex)
               // {
               //     MessageBox.Show("Payroll cannot be generated. Something went wrong. " + ex.Message);
               // }

            }
            else
            {
                MessageBox.Show("No generated DTR is available.");

            }
        }

        private void inc_pbar(int i, int rw)
        {
            try
            {

                if (pbar.Value <= rw)
                {
                    pbar.Invoke(new Action(() =>
                    {
                        pbar.Value = i;
                    }));

                }
                else
                {
                    pbar.Invoke(new Action(() =>
                    {
                        pbar.Value = rw;
                    }));
                }

            }
            catch (Exception)
            {

            }
        }
        public void disp_list()
        {
            DataTable dt = db.QueryBySQLCode("");

        }
        private DataTable get_generated_dtr()
        {
            DataTable dt = null;
            try
            {
                dt = db.QueryBySQLCode("SELECT * from rssys.hr_dtr_sum_employees WHERE isgenerated = '0'");
            }
            catch { }

            return dt;
        }

        private String get_payrol_period(String ppid)
        {
            String period = "";
            DataTable dt = null;
            try
            {
                dt = db.QueryBySQLCode("SELECT concat(to_char(date_from, 'mm/dd/yyyy'),' To ',to_char(date_to, 'mm/dd/yyyy')) as period FROM rssys.hr_payrollpariod WHERE pay_code='" + ppid + "'");
            }
            catch { }
            if(dt.Rows.Count > 0)
            {
                period = dt.Rows[0]["period"].ToString();
            }
            return period;
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
                        query = "SELECT DISTINCT work_date,(SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.work_date = '" + day.ToString("yyyy-MM-dd") + "' AND e.empid ='" +empid +"' ORDER BY work_date ";
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
        public Double get_legal_hol_ot(String empid,String ppid)
        {
            
            DateTime date_from , date_to;
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
                        holiday = db.QueryBySQLCode("SELECT date_holiday,holiday_type FROM rssys.hr_holidays WHERE date_holiday = '" + datein + "' AND holiday_type = 'L'");
                        if (holiday.Rows.Count > 0 )
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

        public DataTable get_holidays(String date_from,String date_to)
        {
            String query = "SELECT * FROM rssys.hr_holidays WHERE date_holiday BETWEEN '" + date_from + "' AND '" + date_to+"'";
            return db.QueryBySQLCode(query);
        }

        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }
        private void p_GeneratePayroll_Load(object sender, EventArgs e)
        {
            disp_list_history();
        }

        
        private void disp_list_history()
        {
            dgv_list.Rows.Clear();
            try
            {

                DataTable dt = db.QueryBySQLCode("SELECT sh.*, concat(firstname,' ',lastname) AS employee FROM rssys.hr_dtr_sum_hdr sh LEFT JOIN rssys.hr_employee e ON e.empid=sh.empid  ORDER BY date_generated DESC, time_generated DESC");

                if (dt.Rows.Count > 0)
                {
                    for (int r = 0; dt.Rows.Count > r; r++)
                    {
                        int i = dgv_list.Rows.Add();
                        DataGridViewRow row = dgv_list.Rows[i];

                        row.Cells["dgvl_date"].Value = gm.toDateString(dt.Rows[r]["date_generated"].ToString(), "");
                        row.Cells["dgvl_time"].Value = dt.Rows[r]["time_generated"].ToString();
                        row.Cells["dgvl_payroll"].Value = gm.toDateString(dt.Rows[r]["date_from"].ToString(), "") + " TO " + gm.toDateString(dt.Rows[r]["date_to"].ToString(), "");

                        row.Cells["dgvl_userid"].Value = dt.Rows[r]["empid"].ToString();
                        //cbo_employee.SelectedValue = dt.Rows[r]["empid"].ToString();
                        row.Cells["dgvl_employee"].Value = dt.Rows[r]["employee"].ToString();

                    }
                }
            }
            catch { }

        }
    }
}
