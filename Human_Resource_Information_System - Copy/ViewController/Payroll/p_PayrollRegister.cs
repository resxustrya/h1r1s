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
    public partial class p_PayrollRegister : Form
    {
        thisDatabase db = new thisDatabase();
        public p_PayrollRegister()
        {
            InitializeComponent();
            disp_list();
        }

        private void btn_new_Click(object sender, EventArgs e)
        {
            p_PayrollRegisterEntry frm = new p_PayrollRegisterEntry(this,"",true);
            frm.ShowDialog();
            disp_list();
        }

        private void p_PayrollRegister_Load(object sender, EventArgs e)
        {

        }

        private void btn_upd_Click(object sender, EventArgs e)
        {
            int r = 0;
            String code;
            if (dgv_list.Rows.Count > 1)
            {
                r = dgv_list.CurrentRow.Index;
                code = dgv_list["dgvl_payrollreg_code", r].Value.ToString();
                p_PayrollRegisterEntry frm = new p_PayrollRegisterEntry(this,code, false);
                frm.ShowDialog();
                disp_list();
            }
            else {

                MessageBox.Show("No records to be selected.");
            }
        }

        public void disp_list()
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();

            dt = db.QueryBySQLCode("SELECT r.payrollreg_code, r.payroll_period_code, r.dept_frm, r.dept_until, r.employee, r.report_type, r.payroll_period_desc, e.lastname || ',' || e.firstname  AS empname FROM rssys.hr_payroll_register r LEFT JOIN rssys.hr_employee e ON r.employee=e.empid ORDER BY payrollreg_code");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_payrollreg_code"].Value = dt.Rows[r]["payrollreg_code"].ToString();
                row.Cells["dgvl_payroll_period_code"].Value = dt.Rows[r]["payroll_period_code"].ToString();
                row.Cells["dgvl_dept_frm"].Value = dt.Rows[r]["dept_frm"].ToString();
                row.Cells["dgvl_dept_until"].Value = dt.Rows[r]["dept_until"].ToString();
                row.Cells["dgvl_employee"].Value = dt.Rows[r]["employee"].ToString();
                row.Cells["dgvl_employee_name"].Value = dt.Rows[r]["empname"].ToString();
                row.Cells["dgvl_report_type"].Value = dt.Rows[r]["report_type"].ToString();
                row.Cells["dgvl_payroll_period_desc"].Value = dt.Rows[r]["payroll_period_desc"].ToString();
               


            }
        }

        private void btn_mainsave_Click(object sender, EventArgs e)
        {

        }
    }
}
