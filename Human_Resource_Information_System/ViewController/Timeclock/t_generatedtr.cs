using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace Human_Resource_Information_System
{
    public partial class t_generatedtr : Form
    {
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();

        Boolean isUseCboEmp = false;
        public t_generatedtr()
        {
            InitializeComponent();
        }

        private void t_generatedtr_Load(object sender, EventArgs e)
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();

            gc.load_employee(cbo_employee);
            gc.load_payroll_period(cbo_payroll_period);

            disp_list_history();
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            if (cbo_payroll_period.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a payroll period.");
                cbo_payroll_period.DroppedDown = true;
                return;
            }

            bgWorker.RunWorkerAsync();
        }

        private String compute_late(String empid, String date_from, String date_to)
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
            return result;
        }

        private String compute_undertime(String empid, String date_from,String date_to)
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

            return result;
        }

        private String compute_overtime(String empid, String date_from, String date_to)
        {
            String result = "00:00:00";

            String query = "";
            String timein = "", timeout = "";
            String time_from = "", time_to = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);

            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            if(sched.Rows.Count > 0)
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
           
            return result;
        }

        private String compute_daysworked(String empid, String date_from, String date_to)
        {
            String result = "0";
            DataTable sched = db.QueryBySQLCode("SELECT shift_sched_from,shift_sched_to FROM rssys.hr_employee WHERE empid = '" + empid + "'");
            String query = "";
            
            if (sched.Rows.Count > 0) {

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
                if (dayno !=  off1 && dayno != off2)
                {
                    total++;
                }
            }
            try
            {
                result = (Convert.ToInt32(total) - Convert.ToInt32(total_worked)).ToString();
               // MessageBox.Show("Total working days :" + total + " Total worked days : " + total_worked + "Result is : " + result);
            }catch(Exception ex)
            {
                result = "0";
            }
            
            return result;
        }
        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }


        public Boolean save_summary(String code)
        {
            Boolean ok = false;
            String table = "hr_dtr_sum_employees", col = "", val = "", empid = "", summ_code = "", days_worked = "", absences = "", late = "", undertime = "", total_overtime = "", summary_code="";
            Boolean success = false;
            int selectedIndx = - 1;

            try
            {

                for (int r = dgv_list_logs.Rows.Count - 1; r >= 0; r--)
                {
                    Boolean use = false;
                    cbo_employee.Invoke(new Action(() =>
                    {
                        empid = dgv_list_logs["empid", r].Value.ToString();

                        if (cbo_employee.SelectedIndex == -1)
                        {
                            use = true;
                        }
                        else
                        {
                            if (empid != cbo_employee.SelectedValue)
                            {
                                use = true;
                                selectedIndx++;
                            }
                        }

                      //  if (use && selectedIndx < 1) 
                      //  {
                            days_worked = dgv_list_logs["days_worked", r].Value.ToString();
                            absences = dgv_list_logs["absent", r].Value.ToString();
                            late = dgv_list_logs["total_late", r].Value.ToString();
                            undertime = dgv_list_logs["undertime", r].Value.ToString();
                            total_overtime = dgv_list_logs["overtime", r].Value.ToString();

                            col = "empid,days_worked,absences,late,undertime,total_overtime,ppid";
                            val = "'" + empid + "','" + days_worked + "','" + absences + "','" + late + "','" + undertime + "','" + total_overtime + "','" + code + "'";

                            db.DeleteOnTable(table, "ppid='" + code + "' AND empid='" + empid + "' AND isgenerated=0");
                            if (!is_generated_sum(empid, code))
                            {
                                if (db.InsertOnTable(table, col, val))
                                {
                                    success = true;
                                    //db.set_pkm99("summ_code", db.get_nextincrementlimitchar(summ_code, 8));
                                }
                            }
                            else
                            {
                                MessageBox.Show("The payroll period for this employee no. " + empid + " is already generated to Payroll System. DTR can not be re-generated.");
                                success = false;
                            }
                       // }

                    }));

                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }
           

            return success;
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

        private void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try { dgv_list_logs.Rows.Clear(); }
            catch (Exception) { }
            String days_to_work = "";
            String empid = "", date_from = "", date_to = "", col = "", val = "", d_from = "", d_to = "", d_now = "", t_now = "", table = "", code = "";
            DataTable dt = null;
            DataTable logs = null;
            DataTable pay_period = null;
            String query = "", total_worked = "", pay_code = "";
            Boolean success = false;
            int j = 0, bar = 1, indx = -1;

            //try { dgv_list.Rows.Clear(); }
            //catch { }


            cbo_payroll_period.Invoke(new Action(() =>
            {
                indx = cbo_employee.SelectedIndex;
                if (cbo_payroll_period.SelectedIndex != -1)
                {
                    pay_code = cbo_payroll_period.SelectedValue.ToString();
                    pay_period = get_date(pay_code);
                    if (pay_period.Rows.Count > 0)
                    {
                        date_from = pay_period.Rows[0]["date_from"].ToString();
                        date_to = pay_period.Rows[0]["date_to"].ToString();
                    }
                }
            }));
            

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
            dt = db.QueryBySQLCode(query);
            if (dt.Rows.Count > 0)
            {
                pbar.Invoke(new Action(() =>
                {
                    pbar.Maximum = dt.Rows.Count;
                }));
                for (int r = 0; r < dt.Rows.Count; r++)
                {

                    empid = dt.Rows[r]["empid"].ToString();
                    dgv_list_logs.Invoke(new Action(() =>
                    {
                        j = dgv_list_logs.Rows.Add();
                        DataGridViewRow row = dgv_list_logs.Rows[j];
                        row.Cells["empid"].Value = dt.Rows[r]["empid"].ToString();
                        row.Cells["name"].Value = dt.Rows[r]["firstname"].ToString() + ", " + dt.Rows[r]["lastname"].ToString();
                        row.Cells["days_worked"].Value = total_worked = compute_daysworked(dt.Rows[r]["empid"].ToString(), date_from, date_to);
                        row.Cells["absent"].Value = compute_absent(dt.Rows[r]["empid"].ToString(), date_from, date_to, total_worked);

                        row.Cells["total_late"].Value = compute_late(dt.Rows[r]["empid"].ToString(), date_from, date_to);
                        row.Cells["undertime"].Value = compute_undertime(dt.Rows[r]["empid"].ToString(), date_from, date_to);
                        row.Cells["overtime"].Value = compute_overtime(dt.Rows[r]["empid"].ToString(), date_from, date_to);
                        j++;
                        inc_pbar(bar, dt.Rows.Count);
                        bar++;
                    }));
                    
                }

                DialogResult result = MessageBox.Show("Do you want to save generated " + (indx == -1 ? "all" : "") + " employee's DTR?", "Confirmation", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    cbo_employee.Invoke(new Action(() =>
                    {
                        String[] empids = new String[1];

                        if (cbo_employee.SelectedIndex != -1)
                        {
                            empids[0] = cbo_employee.SelectedValue.ToString();
                        }
                        else
                        {
                            empids = new String[cbo_employee.Items.Count];
                            for (int i = 0; i < empids.Length; i++)
                            {
                                cbo_employee.SelectedIndex = i;
                                empids[i] = cbo_employee.SelectedValue.ToString();
                            }
                        }

                        foreach (String _empid in empids)
                        {
                            table = "hr_dtr_sum_hdr";
                            code = pay_code;
                            d_from = gm.toDateString(date_from, "");
                            d_to = gm.toDateString(date_to, "");
                            d_now = DateTime.Now.ToString("yyyy-MM-dd");
                            t_now = DateTime.Now.ToString("HH:mm");
                            //col = "ppid,date_from,date_to,date_generated";
                            //val = "'" + code + "','" + d_from + "','" + d_to + "','" + d_now + "'";
                            col = "empid,ppid,date_from,date_to,date_generated,time_generated";
                            val = "'" + _empid + "','" + code + "','" + d_from + "','" + d_to + "','" + d_now + "','" + t_now + "'";

                            db.InsertOnTable(table, col, val);
                        }

                    }));


                    success = true;
                    if (save_summary(code))
                    {
                        // try { dgv_list_logs.Rows.Clear(); }
                        // catch (Exception) { }
                        cbo_employee.Invoke(new Action(() =>
                        {
                            cbo_employee.SelectedIndex = -1;
                            disp_list_history();
                        }));
                        cbo_payroll_period.Invoke(new Action(() =>
                        {
                            cbo_payroll_period.SelectedIndex = -1;
                        }));
                        
                        MessageBox.Show("New DTR summary saved.");
                        pbar.Invoke(new Action(() =>
                        {
                            pbar.Value = 0;
                        }));
                    }
                    else
                    {
                        success = false;
                        db.DeleteOnTable(table, "summary_code='" + code + "'");
                        MessageBox.Show("Failed on saving.");
                    }

                }

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

        private Boolean is_generated_sum(String empid, String code)
        {
            Boolean ok = false;
            try
            {
                DataTable dt = db.QueryBySQLCode("SELECT empid FROM rssys.hr_dtr_sum_employees WHERE empid = '" + empid + "' AND ppid ='" + code + "'");

                if (dt.Rows.Count > 0)
                {
                    ok = true;
                }
            }
            catch { }
           
            return ok;
        }

        private void cbo_employee_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUseCboEmp == false) 
                disp_list_history();
        }
        private void cbo_payroll_period_SelectedIndexChanged(object sender, EventArgs e)
        {
            disp_list_history();
        }

        private void disp_list_history()
        {
            dgv_list.Rows.Clear();
            try
            {
                String WHERE = "WHERE";

                if (cbo_employee.SelectedIndex != -1)
                {
                    WHERE += " empid='" + cbo_employee.SelectedValue + "'";
                }
                if (cbo_payroll_period.SelectedIndex != -1)
                {
                    WHERE += (WHERE == "WHERE" ? "" : " AND ");
                    WHERE += " ppid='" + cbo_payroll_period.SelectedValue + "'";
                }

                if (WHERE == "WHERE")
                    WHERE = "";

                DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_dtr_sum_hdr " + WHERE + "  ORDER BY date_generated DESC, time_generated DESC");

                if (dt.Rows.Count > 0)
                {
                    int indx = cbo_employee.SelectedIndex;
                    for (int r = 0; dt.Rows.Count > r; r++) {
                        int i = dgv_list.Rows.Add();
                        DataGridViewRow row = dgv_list.Rows[i];

                        row.Cells["dgvl_date"].Value = gm.toDateString(dt.Rows[r]["date_generated"].ToString(), "");
                        row.Cells["dgvl_time"].Value = dt.Rows[r]["time_generated"].ToString();
                        row.Cells["dgvl_payroll"].Value = gm.toDateString(dt.Rows[r]["date_from"].ToString(), "") + " TO " + gm.toDateString(dt.Rows[r]["date_to"].ToString(), "");

                        row.Cells["dgvl_userid"].Value = dt.Rows[r]["empid"].ToString();
                        isUseCboEmp = true;
                        cbo_employee.SelectedValue = dt.Rows[r]["empid"].ToString();
                        row.Cells["dgvl_employee"].Value = cbo_employee.Text;

                    }
                    cbo_employee.SelectedIndex = indx;
                    isUseCboEmp = false;
                }
            }
            catch { }
            
        }

        private void dgv_list_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
             
        }



        
    }
}
