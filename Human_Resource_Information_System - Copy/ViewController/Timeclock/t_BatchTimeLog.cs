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
    public partial class t_BatchTimeLog : Form
    {

        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();
        t_BatchTimeLog bt;

        private String old_date = "", old_logs = "";


        Boolean update_log = false, isBtnGo = false;
        String isEmp = "";

        public t_BatchTimeLog()
        {
            InitializeComponent();

            dtp_frm.Value = DateTime.Parse(dtp_frm.Value.ToString("yyyy-MM-01"));
        }

        private void btn_itemadd_Click(object sender, EventArgs e)
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

            update_log = false;
            isnew = true;
            t_addLogs add = new t_addLogs(this);
            add.ShowDialog();
            
        }

        private void t_BatchTimeLog_Load(object sender, EventArgs e)
        {
            gm = new GlobalMethod();
            gc = new GlobalClass();

            gc.load_employee(cbo_employee);
        }
        public void set_logs(String work_date,String time_logs,String status,Boolean new_logs)
        {
            int i = 0;
            if (update_log == false)
            {
                i = dgv_list.Rows.Add();
            }
            else {
                i = dgv_list.CurrentRow.Index;
            }

            DataGridViewRow row = dgv_list.Rows[i];

            row.Cells["work_date"].Value = gm.toDateString(work_date, "");
            row.Cells["time_log"].Value = gm.toDateString(time_logs, "HH:mm");
            row.Cells["is_new"].Value = new_logs;
            if (status == "I")
            {
                row.Cells["status"].Value = "IN";
            }
            else
            {
                row.Cells["status"].Value = "OUT";
            }
            
        }

        private void btn_mainsave_Click(object sender, EventArgs e)
        {
            String col = "", val = "";
            Boolean success = false;
            String table = "hr_tito2";
            String empid = "", work_date = "", time_log = "", status = "", source = "";

            String date_from = dtp_frm.Value.ToString("yyyy-MM-dd");
            String date_to = dtp_to.Value.ToString("yyyy-MM-dd");


            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }
            if (dgv_list.Rows.Count <= 1)
            {
                MessageBox.Show("Empty batch time list.");
                return;
            }
            if (isEmp != cbo_employee.SelectedValue.ToString() && isBtnGo)
            {
                MessageBox.Show("Invalid selection of employee.");
                return;
            }

            empid = cbo_employee.SelectedValue.ToString();
            source = "M";

            if (isBtnGo == true)
            {
                db.UpdateOnTable(table, "cancel='Y'", "work_date BETWEEN '" + date_from + "' AND '" + date_to + "' AND empid ='" + empid + "'");
            }

            for (int r = 0; r < dgv_list.Rows.Count - 1; r++)
            {
                String new_logs = "";

                try { new_logs = dgv_list["is_new", r].Value.ToString(); }
                catch { }


                work_date = dgv_list["work_date", r].Value.ToString();
                DateTime dt = DateTime.Parse(dgv_list["time_log", r].Value.ToString());
                time_log = dt.ToString("HH:mm");
                status = dgv_list["status", r].Value.ToString().Substring(0, 1);

                if (new_logs == "True")
                {
                    col = "work_date,time_log,empid,status,source";
                    val = "'" + work_date + "','" + time_log + "','" + empid + "','" + status + "','" + source + "'";

                    db.InsertOnTable(table, col, val);

                    success = true;
                }
                else
                {
                    db.UpdateOnTable(table, "cancel=''", "empid='" + empid + "' AND work_date='" + work_date + "' AND time_log ='" + time_log + "'");
                }
            }

            if (isBtnGo == true)
            {
                db.DeleteOnTable(table, "cancel='Y'");
            }

            if (success == true)
            {
                MessageBox.Show("New time logs was saved for Employee ID : " + empid);

                old_logs = "";
                old_date = "";
                update_log = false;
                isBtnGo = true;
                disp_logs();
            }
            else
            {
                MessageBox.Show("Some time logs was removed for Employee ID : " + empid);
                disp_logs();
            }
            clear();
        }
        private void disp_logs()
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

                    DataTable dt = db.QueryBySQLCode("SELECT to_char(work_date, 'yyyy-MM-dd') AS work_date, time_log, status, source FROM rssys.hr_tito2 WHERE work_date BETWEEN '" + date_from + "' AND '" + date_to + "' AND empid ='" + empid + "' ORDER BY work_date ASC");

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

       
        private void button1_Click(object sender, EventArgs e) // GO BUTTON
        {

            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }

            isEmp = cbo_employee.SelectedIndex != -1 ? cbo_employee.SelectedValue.ToString() : "";
            isBtnGo = true;
            disp_logs();
        }

        private void btn_itemupd_Click(object sender, EventArgs e)
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
                String status = dgv_list["status", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(status))
                {
                    update_log = true;

                    try
                    {
                        String work_date = dgv_list["work_date", r].Value.ToString();
                        String time_log = dgv_list["time_log", r].Value.ToString();

                        old_date = work_date;
                        old_logs = time_log;

                        t_addLogs add = new t_addLogs(this);
                        add.set_log(work_date, time_log, status);
                        add.ShowDialog();
                       
                    }
                    catch { }

                }
                else
                {
                    MessageBox.Show("No batch time selected.");
                }
            }
            catch 
            {
                MessageBox.Show("No batch time selected.");
            }
        }

        private void cbo_employee_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void btn_itemremove_Click(object sender, EventArgs e)
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
                String code = dgv_list["status", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(code))
                {
                    DialogResult result = MessageBox.Show("Are you sure you want to cancel this time log?", "", MessageBoxButtons.YesNoCancel);
                    if (result == DialogResult.Yes)
                    {
                        dgv_list.Rows.RemoveAt(r);
                    }
                }
                else
                {
                    MessageBox.Show("No batch time selected.");
                }
            }
            catch
            {
                MessageBox.Show("No batch time selected.");
            }
        }

        private void btn_mainexit_Click(object sender, EventArgs e)
        {
            clear();
        }
        void clear()
        {
            cbo_employee.SelectedIndex = -1;
            dgv_list.Rows.Clear();
        }
    }
}
