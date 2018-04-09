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
    public partial class t_manpowerschedule : Form
    {

        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
        t_BatchTimeLog bt;

        private String old_date = "", old_logs = "";


        Boolean isBtnGo = false;
        String isEmp = "";

        public t_manpowerschedule()
        {
            InitializeComponent();

            dtp_frm.Value = DateTime.Parse(dtp_frm.Value.ToString("yyyy-MM-01"));
        }
        private void t_manpowerschedule_Load(object sender, EventArgs e)
        {
            gc.load_employee(cbo_employee);
        }

        private void btn_addwork_Click(object sender, EventArgs e)
        {

            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }
            if (isEmp != cbo_employee.SelectedValue.ToString() && isBtnGo)
            {
                MessageBox.Show("Invalid selection of employee.");
                return;
            }

            isnew = true;

            t_addSched add = new t_addSched(this,true);
            add.ShowDialog();
            
        }

        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }

        public void set_work(Boolean _isnew,String wrkd_frm, String wrkd_to, String timein1, String timein2, String timeout1, String timeout2, String shift_sched_id, String shift_sched_desc)
        {
            if (!_isnew) {
                dgv_list.Rows.RemoveAt(dgv_list.CurrentRow.Index);
            }

            DateTime StartDate = DateTime.Parse(wrkd_frm);
            DateTime EndDate = DateTime.Parse(wrkd_to);

            foreach (DateTime day in EachDay(StartDate, EndDate))
            {
                int r = dgv_list.Rows.Add();

                // dgvl_restday dgvl_otin  dgvl_otout
                dgv_list["dgvl_date", r].Value = day.ToString("yyyy-MM-dd");
                dgv_list["dgvl_day", r].Value = day.ToString("dddd");
                dgv_list["dgvl_dayn", r].Value = day.DayOfWeek.ToString("D");
                dgv_list["dgvl_timein1", r].Value = timein1;
                dgv_list["dgvl_timein2", r].Value = timein2;
                dgv_list["dgvl_timeout1", r].Value = timeout1;
                dgv_list["dgvl_timeout2", r].Value = timeout2;
                dgv_list["dgvl_shift_sched", r].Value = shift_sched_desc;
                dgv_list["dgvl_shiftsched_code", r].Value = shift_sched_id;
                dgv_list["is_new", r].Value = "True";

            }
        }

        private void btn_mainsavework_Click(object sender, EventArgs e)
        {
            String col = "", val = "";
            Boolean success = false;
            String table = "hr_work_schedule";


            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }
            if (isEmp != cbo_employee.SelectedValue.ToString() && isBtnGo)
            {
                MessageBox.Show("Invalid selection of employee.");
                return;
            }
            if (dgv_list.Rows.Count <= 1)
            {
                MessageBox.Show("Empty schedule list.");
                return;
            }


            String date_from = dtp_frm.Value.ToString("yyyy-MM-dd");
            String date_to = dtp_to.Value.ToString("yyyy-MM-dd");
            String empid = cbo_employee.SelectedValue.ToString();
          

            if (isBtnGo == true)
            {
                db.UpdateOnTable(table, "cancel='Y'", "wrk_date BETWEEN '" + date_from + "' AND '" + date_to + "' AND empid ='" + empid + "'");
            }

            for (int r = 0; r < dgv_list.Rows.Count - 1; r++)
            {
                String new_logs = "";
                try { new_logs = dgv_list["is_new", r].Value.ToString(); }
                catch { }

                String wrk_date = dgv_list["dgvl_date", r].Value.ToString();
                String wrk_day = dgv_list["dgvl_day", r].Value.ToString();
                String wrk_dayn = dgv_list["dgvl_dayn", r].Value.ToString();
                String shift_sched = dgv_list["dgvl_shift_sched", r].Value.ToString();
                String shiftsched_code = dgv_list["dgvl_shiftsched_code", r].Value.ToString();
                String timein1 = dgv_list["dgvl_timein1", r].Value.ToString();
                String timein2 = dgv_list["dgvl_timein2", r].Value.ToString();
                String timeout1 = dgv_list["dgvl_timeout1", r].Value.ToString();
                String timeout2 = dgv_list["dgvl_timeout2", r].Value.ToString();
                String restday = (dgv_list["dgvl_restday", r].Value ?? "").ToString();
                String ot_in = (dgv_list["dgvl_otin", r].Value ?? "").ToString();
                String ot_out = (dgv_list["dgvl_otout", r].Value ?? "").ToString();

                //empid, wrk_date, wrk_day, wrk_dayn, shift_sched, shiftsched_code, timein1, timein2, timeout1, timeout2, restday, ot_in, ot_out

                if (new_logs == "True")
                {
                    col = "empid, wrk_date, wrk_day, wrk_dayn, shift_sched, shiftsched_code, timein1, timein2, timeout1, timeout2, restday, ot_in, ot_out";
                    val = "'" + empid + "','" + wrk_date + "', '" + wrk_day + "', '" + wrk_dayn + "', '" + shift_sched + "', '" + shiftsched_code + "', '" + timein1 + "', '" + timein2 + "', '" + timeout1 + "', '" + timeout2 + "', '" + restday + "', '" + ot_in + "', '" + ot_out + "'";

                    db.InsertOnTable(table, col, val);

                    success = true;
                }
                else
                {
                    String code = dgv_list["dgvl_code", r].Value.ToString();
                    db.UpdateOnTable(table, "cancel=''", "code = '" + code + "'");
                }
            }

            if (isBtnGo == true)
            {
                db.DeleteOnTable(table, "cancel='Y'");
            }

            if (success == true)
            {
                MessageBox.Show("New schedule was saved for Employee ID : " + empid);

                isBtnGo = true;
                disp_schedule();
            }
            else
            {
                MessageBox.Show("Some schedule was removed for Employee ID : " + empid);
                disp_schedule();
            }
        }
        private void disp_schedule()
        {
            dgv_list.Rows.Clear();

            int i = 0;
            try
            {
                if (cbo_employee.SelectedIndex != -1)
                {
                    String empid = cbo_employee.SelectedValue.ToString();
                    String date_from = dtp_frm.Value.ToString("yyyy-MM-dd");
                    String date_to = dtp_to.Value.ToString("yyyy-MM-dd");

                    DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_work_schedule WHERE wrk_date BETWEEN '" + date_from + "' AND '" + date_to + "' AND empid ='" + empid + "' ORDER BY wrk_date ASC");
                    if (dt != null)
                    {
                        for (int r = 0; r < dt.Rows.Count; r++)
                        {
                            i = dgv_list.Rows.Add();
                            DataGridViewRow row = dgv_list.Rows[i];

                            row.Cells["dgvl_code"].Value = dt.Rows[r]["code"].ToString();

                            row.Cells["dgvl_date"].Value = gm.toDateString(dt.Rows[r]["wrk_date"].ToString(), "");
                            row.Cells["dgvl_day"].Value = dt.Rows[r]["wrk_day"].ToString();
                            row.Cells["dgvl_dayn"].Value = dt.Rows[r]["wrk_dayn"].ToString();
                            row.Cells["dgvl_shift_sched"].Value = dt.Rows[r]["shift_sched"].ToString();
                            row.Cells["dgvl_shiftsched_code"].Value = dt.Rows[r]["shiftsched_code"].ToString();
                            row.Cells["dgvl_timein1"].Value = dt.Rows[r]["timein1"].ToString();
                            row.Cells["dgvl_timein2"].Value = dt.Rows[r]["timein2"].ToString();
                            row.Cells["dgvl_timeout1"].Value = dt.Rows[r]["timeout1"].ToString();
                            row.Cells["dgvl_timeout2"].Value = dt.Rows[r]["timeout2"].ToString();
                            row.Cells["dgvl_restday"].Value = dt.Rows[r]["restday"].ToString();
                            row.Cells["dgvl_otin"].Value = dt.Rows[r]["ot_in"].ToString();
                            row.Cells["dgvl_otout"].Value = dt.Rows[r]["ot_out"].ToString();

                        }
                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show("Empty schedule.");
                        }
                    }
                }
            }
            catch { MessageBox.Show("Please select an employee."); }

            /*
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }

            int i = 0;
            String date_from = dtp_frm.Value.ToString("yyyy-MM-dd");
            String date_to = dtp_to.Value.ToString("yyyy-MM-dd");

            String empid = cbo_employee.SelectedValue.ToString();

            DataTable dt = db.QueryBySQLCode("SELECT to_char(work_date, 'dd/MM/yyyy') AS work_date, time_log, status, source FROM rssys.hr_tito2 WHERE work_date BETWEEN '" + date_from + "' AND '" + date_to + "' AND empid ='" + empid + "'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["work_date"].Value = dt.Rows[r]["work_date"].ToString();
                row.Cells["time_log"].Value = dt.Rows[r]["time_log"].ToString();

                if (dt.Rows[r]["status"].ToString() == "I")
                {
                    row.Cells["status"].Value = "IN";
                }
                else
                {
                    row.Cells["status"].Value = "OUT";
                }

                if (dt.Rows[r]["source"].ToString() == "M")
                {
                    row.Cells["source"].Value = "Manual";
                }
                i++;
            }*/
        }


        private void btn_filter_Click(object sender, EventArgs e) // GO BUTTON
        {

            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }

            isEmp = (cbo_employee.SelectedValue??"").ToString();
            isBtnGo = true;
            disp_schedule();
        }

        private void btn_updwork_Click(object sender, EventArgs e)
        {
            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }
            if (isEmp != cbo_employee.SelectedValue.ToString() && isBtnGo)
            {
                MessageBox.Show("Invalid selection of employee.");
                return;
            }

            try
            {
                int r = dgv_list.CurrentRow.Index;
                String status = dgv_list["dgvl_date", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(status))
                {
                    try
                    {
                        t_addSched add = new t_addSched(this, false);
                        add.ShowDialog();
                    }
                    catch { }
                }
                else
                {
                    MessageBox.Show("No schedule selected.");
                }
            }
            catch 
            {
                MessageBox.Show("No schedule selected.");
            }
        }


        private void btn_delwork_Click(object sender, EventArgs e)
        {
            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }
            if (isEmp != cbo_employee.SelectedValue.ToString() && isBtnGo)
            {
                MessageBox.Show("Invalid selection of employee.");
                return;
            }

            try
            {
                int r = dgv_list.CurrentRow.Index; 
                String str = dgv_list["dgvl_date", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(str))
                {

                    DialogResult result = MessageBox.Show("Are you sure you want to cancel this schedule?", "", MessageBoxButtons.YesNoCancel);
                    if (result == DialogResult.Yes)
                    {
                        dgv_list.Rows.RemoveAt(r);
                    }
                }
                else
                {
                    MessageBox.Show("No schedule selected.");
                }
            }
            catch
            {
                MessageBox.Show("No schedule selected.");
            }
        }

        public DataGridView getDGV() {
            return dgv_list;
        }



    }
}
