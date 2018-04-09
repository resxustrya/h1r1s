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
    public partial class t_emp_dtr : Form
    {
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();
   
        public t_emp_dtr()
        {
            InitializeComponent();
        }

        private void t_emp_dtr_Load(object sender, EventArgs e)
        {
            gc.load_employee(cbo_employee);

            disp_list(false);
        }



        private void btn_save_Click(object sender, EventArgs e)
        {
            disp_list(true);
        }

        private void disp_list(Boolean withMsg)
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;
            String empid = "";
            String query = "";
            String WHERE = "";

            if (cbo_employee.SelectedIndex != -1)
            {
                empid = cbo_employee.SelectedValue.ToString();
                WHERE = "AND t.empid ='" + empid + "'";
            }

            String date_from = dtp_frm.Value.ToString("yyyy-MM-dd");
            String date_to = dtp_to.Value.ToString("yyyy-MM-dd");

            query = "SELECT DISTINCT CONCAT(lastname,' ',firstname) AS name, t.source, e.empid, to_char(work_date, 'yyyy-MM-dd') AS work_date, (SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.work_date BETWEEN '" + date_from + "' AND '" + date_to + "' " + WHERE + " ORDER BY work_date ";

            DataTable dt = db.QueryBySQLCode(query);

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["name"].Value = dt.Rows[i]["name"].ToString();
                row.Cells["work_date"].Value = dt.Rows[i]["work_date"].ToString();
                row.Cells["timein"].Value = dt.Rows[i]["timein"].ToString();
                row.Cells["timeout"].Value = dt.Rows[i]["timeout"].ToString();

                if (dt.Rows[i]["source"].ToString() == "M")
                {
                    row.Cells["source"].Value = "Manual";
                }
            }

            if (dt.Rows.Count == 0 && withMsg)
            {
                if (empid == "")
                {
                    MessageBox.Show("Has no employee's DTR at this date.");
                }
                else
                {
                    MessageBox.Show("No DTR for this employee.");
                }
            }
        }

        private void cbo_employee_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }
    }
}
