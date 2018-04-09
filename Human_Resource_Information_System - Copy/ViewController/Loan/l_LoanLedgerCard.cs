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
    public partial class l_LoanLedgerCard : Form
    { thisDatabase db = new thisDatabase();
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        Boolean isnew = false;
        Boolean seltbp = false;
        public l_LoanLedgerCard()
        {
            InitializeComponent();
            gc.load_employee(cbo_employee);
        }

        private void l_LoanLedgerCard_Load(object sender, EventArgs e)
        {

        }
        public void disp_list(String empcode)
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();

            dt = db.QueryBySQLCode("SELECT loan_code, loan_desc,loan_transdate, loan_location, loan_type, user_id, whs_location_code, loan_cost_center_code,loan_cost_center_name,loan_sub_cost_center,employee_no,employee_name,loan_amount,loan_deduction FROM rssys.hr_loanhdr  WHERE employee_no='" + empcode + "'ORDER BY loan_code");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["loan_code"].ToString();
                row.Cells["dgvl_description"].Value = dt.Rows[r]["loan_desc"].ToString();
                row.Cells["dgvl_location"].Value = dt.Rows[r]["loan_location"].ToString();
                row.Cells["dgvl_trn_type"].Value = dt.Rows[r]["loan_type"].ToString();
                row.Cells["dgvl_userid"].Value = dt.Rows[r]["user_id"].ToString();
                row.Cells["dgvl_whscode"].Value = dt.Rows[r]["whs_location_code"].ToString();
                row.Cells["dgvl_ccname"].Value = dt.Rows[r]["loan_cost_center_name"].ToString();
                row.Cells["dgvl_sub_costcenter"].Value = dt.Rows[r]["loan_sub_cost_center"].ToString();
                row.Cells["dgvl_empno"].Value = dt.Rows[r]["employee_no"].ToString();
                row.Cells["dgvl_empname"].Value = dt.Rows[r]["employee_name"].ToString();
                row.Cells["dgvl_loan_amount"].Value = dt.Rows[r]["loan_amount"].ToString();
                row.Cells["dgvl_deduction_amount"].Value = dt.Rows[r]["loan_deduction"].ToString();
                row.Cells["dgvl_trnxdate"].Value = dt.Rows[r]["loan_transdate"].ToString();

            }
        }
        private void goto_tbcntrl_list()
        {
            seltbp = true;
            tbcntrl_main.SelectedTab = tpg_list;
            tbcntrl_option.SelectedTab = tpg_opt_1;

            tpg_list.Show();
            tpg_opt_1.Show();
            seltbp = false;
        }
        private void btn_new_Click(object sender, EventArgs e)
        {
        //{
        //    isnew = true;
        //    frm_clear();
        //    tpg_info_enable(true);
        //    goto_tbcntrl_info();
        }
        private void tpg_info_enable(Boolean flag)
        {
            //txt_code.Enabled = flag;
        }
        

       

        private void btn_mainexit_Click(object sender, EventArgs e)
        {
            goto_tbcntrl_list();
            tpg_info_enable(false);
            //frm_clear();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {

        }

        private void btn_upd_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_mainsave_Click(object sender, EventArgs e)
        {
           
        }

        private void l_LoanLedgerCard_Load_1(object sender, EventArgs e)
        {

        }

        private void cbo_employee_SelectedIndexChanged(object sender, EventArgs e)
        {
            disp_list(cbo_employee.SelectedValue.ToString());
        }
    }
}
