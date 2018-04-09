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
    public partial class p_OtherDeductionEntry : Form
    {
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
        Boolean yah = false;
        public p_OtherDeductionEntry()
        {
            InitializeComponent();
            gc.load_payroll_period(cbo_payroll_period);
            gc.load_other_deductions(cbo_earnings_code);

            try { cbo_earnings_code.SelectedIndex = 0; }
            catch { cbo_earnings_code.SelectedIndex = -1; }

            try { cbo_payroll_period.SelectedIndex = 0; }
            catch { cbo_payroll_period.SelectedIndex = -1; }

            yah = true;
            disp_list();
        }

        private void btn_new_Click(object sender, EventArgs e)
        {
            p_AddOtherDeductionsEntry frm = new p_AddOtherDeductionsEntry(this, "", true);
            frm.ShowDialog();
            disp_list();
        }
        public void disp_list()
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();

            dt = db.QueryBySQLCode("SELECT de.*,od.description AS deduction FROM rssys.hr_deduction_entry  de LEFT JOIN rssys.hr_other_deductions od ON de.deduction_code=od.code ORDER BY de.dedcode ASC");
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["dedcode"].ToString();
                row.Cells["dgvl_emp_no"].Value = dt.Rows[r]["emp_no"].ToString();
                row.Cells["dgvl_emp_name"].Value = dt.Rows[r]["emp_name"].ToString();
                row.Cells["dgvl_amount"].Value = dt.Rows[r]["amount"].ToString();
                row.Cells["dgvl_payroll_code"].Value = dt.Rows[r]["payroll_period"].ToString();
                row.Cells["dgvl_deduction"].Value = dt.Rows[r]["deduction"].ToString();
            }
        }

        public void disp_list2()
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();
            dt = db.QueryBySQLCode("SELECT de.*,od.description AS deduction FROM rssys.hr_deduction_entry  de LEFT JOIN rssys.hr_other_deductions od ON de.deduction_code=od.code WHERE payroll_period='" + cbo_payroll_period.SelectedValue.ToString() + "' AND deduction_code='" + cbo_earnings_code.SelectedValue.ToString() + "' ORDER BY de.dedcode");
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["dedcode"].ToString();
                row.Cells["dgvl_emp_no"].Value = dt.Rows[r]["emp_no"].ToString();
                row.Cells["dgvl_emp_name"].Value = dt.Rows[r]["emp_name"].ToString();
                row.Cells["dgvl_amount"].Value = dt.Rows[r]["amount"].ToString();
                row.Cells["dgvl_payroll_code"].Value = dt.Rows[r]["payroll_period"].ToString();
                row.Cells["dgvl_deduction"].Value = dt.Rows[r]["deduction"].ToString();
            }
        }

        private void btn_upd_Click(object sender, EventArgs e)
        {
            int r = 0;
            String code;
            if (dgv_list.Rows.Count > 1)
            {
                r = dgv_list.CurrentRow.Index;
                code = dgv_list["dgvl_code", r].Value.ToString();
                p_AddOtherDeductionsEntry frm = new p_AddOtherDeductionsEntry(this, code, false);
                frm.ShowDialog();
                disp_list();
            }
            else
            {

                MessageBox.Show("No records to be selected.");
            }
        }

        private void cbo_payroll_period_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (yah)
                disp_list2();
        }

        private void cbo_earnings_code_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (yah)
                disp_list2();
        }

        private void p_OtherDeductionEntry_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            disp_list();
        }
    }
}
