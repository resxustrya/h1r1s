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
    public partial class p_OtherEarningEntry : Form
    {
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
        Boolean yah = false;
        public p_OtherEarningEntry()
        {
            InitializeComponent();
            gc.load_payroll_period(cbo_payroll_period);
            gc.load_other_earnings(cbo_earnings_code);

            try {cbo_earnings_code.SelectedIndex = 0;}
            catch { cbo_earnings_code.SelectedIndex = -1;  }

            try { cbo_payroll_period.SelectedIndex = 0; }
            catch { cbo_payroll_period.SelectedIndex = -1; }

            yah = true;
            disp_list();
        }

        private void btn_new_Click(object sender, EventArgs e)
        {
            p_AddOtherEarningsEntry frm = new p_AddOtherEarningsEntry(this, "", true);
            frm.ShowDialog();
            disp_list();
        }
        public void disp_list()
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();

            dt = db.QueryBySQLCode("SELECT ee.*,oe.description AS earning FROM rssys.hr_earning_entry ee LEFT JOIN rssys.hr_other_earnings oe ON ee.earning_code=oe.code ORDER BY ee.entcode ASC");
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["entcode"].ToString();
                row.Cells["dgvl_emp_no"].Value = dt.Rows[r]["emp_no"].ToString();
                row.Cells["dgvl_emp_name"].Value = dt.Rows[r]["emp_name"].ToString();
                row.Cells["dgvl_amount"].Value = dt.Rows[r]["amount"].ToString();
                row.Cells["dgvl_payroll_code"].Value = dt.Rows[r]["payroll_period"].ToString();
                row.Cells["dgvl_earning"].Value = dt.Rows[r]["earning"].ToString();
            }
        }

        public void disp_list2()
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();
            dt = db.QueryBySQLCode("SELECT ee.*,oe.description AS earning FROM rssys.hr_earning_entry ee LEFT JOIN rssys.hr_other_earnings oe ON ee.earning_code=oe.code WHERE payroll_period='" + cbo_payroll_period.SelectedValue.ToString() + "' AND earning_code='" + cbo_earnings_code.SelectedValue.ToString() + "' ORDER BY ee.entcode ASC");
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["entcode"].ToString();
                row.Cells["dgvl_emp_no"].Value = dt.Rows[r]["emp_no"].ToString();
                row.Cells["dgvl_emp_name"].Value = dt.Rows[r]["emp_name"].ToString();
                row.Cells["dgvl_amount"].Value = dt.Rows[r]["amount"].ToString();
                row.Cells["dgvl_payroll_code"].Value = dt.Rows[r]["payroll_period"].ToString();
                row.Cells["dgvl_earning"].Value = dt.Rows[r]["earning"].ToString();
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
                p_AddOtherEarningsEntry frm = new p_AddOtherEarningsEntry(this, code, false);
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
            if(yah)
           disp_list2();
        }

        private void cbo_earnings_code_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(yah)
            disp_list2();
        }

        private void p_OtherEarningEntry_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            disp_list();
        }

        private void btn_mainsave_Click(object sender, EventArgs e)
        {

        }
    }
}
