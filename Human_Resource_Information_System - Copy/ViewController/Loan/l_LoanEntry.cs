using System;
using System.Windows;
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
    public partial class l_LoanEntry : Form
    {
        thisDatabase db = new thisDatabase();
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        Boolean isnew = false;
        Boolean seltbp = false;
        public l_LoanEntry()
        {
            InitializeComponent();
            gc.load_employee(cbo_employee);
            gc.load_loantype(cbo_contraacct);
            disp_list();
        }

        private void l_LoanEntry_Load(object sender, EventArgs e)
        {

        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {

        }

        private void btn_itemadd_Click(object sender, EventArgs e)
        {
           
        
                                                                    
        }

        private void btn_mainsave_Click(object sender, EventArgs e)
        {
            Boolean success = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();
            String loan_code, loan_desc, loan_transdate, loan_location, loan_type, user_id, whs_location_code, loan_amount, deduction_amount, loan_cost_center_code, loan_sub_cost_center, employee_code, employee_name, deduction_date;
            String cashier = "";
            String pending = "Y";
            String col = "", val = "", col2 = "", val2 = "";
            String notifyadd = null;
            String table = "hr_loanhdr";
            String tableln = "hr_loanlne";

            //txt_code
            //    txt_desc
            //    cbo_contraacct                                      
            //        cbo_stocklocation                               
            //        cbo_costcenter                                  
            //            cbo_scc     

            if (cbo_contraacct.SelectedIndex == -1)
            {
                MessageBox.Show("No Contract Account Selected");
                cbo_contraacct.DroppedDown = true;
            }
            else if (cbo_stocklocation.SelectedIndex == -1)
            {
                MessageBox.Show("No Stock Location Selected");
                cbo_stocklocation.DroppedDown = true;
            }
            else if(cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("No Employee Selected");
                cbo_employee.DroppedDown = true;
            }
            else if (cbo_costcenter.SelectedIndex == -1)
            {
                MessageBox.Show("No Cost Center Selected");
                cbo_costcenter.DroppedDown = true;
            }
            else if (cbo_scc.SelectedIndex == -1)
            {
                MessageBox.Show("No Sub Cost Center Selected");
                cbo_scc.DroppedDown = true;
            }
            else
            {
                loan_code = txt_code.Text;
                loan_desc = txt_desc.Text;
                loan_location = cbo_stocklocation.Text;
                loan_type = cbo_contraacct.SelectedValue.ToString();
                user_id = GlobalClass.username;
                whs_location_code = cbo_stocklocation.Text;
                loan_cost_center_code = cbo_costcenter.Text;
                loan_sub_cost_center = cbo_scc.Text;
                loan_amount = txt_amnt_loan.Text;
                deduction_amount = txt_deduction.Text;
                employee_code = cbo_employee.SelectedValue.ToString();
                employee_name = cbo_employee.Text;
                loan_transdate = dtp_trnxdt.Value.ToString("yyyy-MM-dd");
                deduction_date = dtp_deduction.Value.ToString("yyyy-MM-dd");
                if (isnew)
                {
                    try
                    {
                        loan_code = db.get_pk("loan_code");
                        col = "loan_code, loan_desc, loan_location, loan_type, user_id, whs_location_code, loan_cost_center_code,loan_cost_center_name,loan_sub_cost_center,employee_no,employee_name,loan_amount,loan_deduction,loan_transdate,deduction_date";
                        val = "" + db.str_E(loan_code) + ", " + db.str_E(loan_desc) + ", " + db.str_E(loan_location) + ", " + db.str_E(loan_type) + ", " + db.str_E(user_id) + ", " + db.str_E(whs_location_code) + ", " + db.str_E(loan_cost_center_code) + ", " + db.str_E(loan_cost_center_code) + ", " + db.str_E(loan_sub_cost_center) + ", " + db.str_E(employee_code) + ", " + db.str_E(employee_name) + ", " + db.str_E(loan_amount) + ", " + db.str_E(deduction_amount) + ", " + db.str_E(loan_transdate) + ", " + db.str_E(deduction_date) + "";

                        if (db.InsertOnTable(table, col, val))
                        {
                            db.set_pkm99("loan_code", db.get_nextincrementlimitchar(loan_code, 8));
                            success = true;
                            //add_items(loan_code);
                        }
                        else
                        {
                            db.DeleteOnTable(table, "loan_code=" + db.str_E(loan_code) + "");
                            MessageBox.Show("Failed on saving.");
                        }
                    }
                    catch { }
                }
                else
                {
                    col = "loan_code=" + db.str_E(loan_code) + ", loan_desc=" + db.str_E(loan_desc) + ", loan_location=" + db.str_E(loan_location) + ", loan_type=" + db.str_E(loan_type) + ", user_id=" + db.str_E(user_id) + ", whs_location_code=" + db.str_E(whs_location_code) + ", loan_cost_center_code=" + db.str_E(loan_cost_center_code) + ", loan_cost_center_name=" + db.str_E(loan_cost_center_code) + ", loan_sub_cost_center=" + db.str_E(loan_sub_cost_center) + ", employee_no=" + db.str_E(employee_code) + ", employee_name=" + db.str_E(employee_name) + ", loan_amount=" + db.str_E(loan_amount) + ", loan_deduction=" + db.str_E(deduction_amount) + ", loan_transdate=" + db.str_E(loan_transdate) + ", deduction_date=" + db.str_E(deduction_date) + "";

                    if (db.UpdateOnTable(table, col, "loan_code=" + db.str_E(loan_code) + ""))
                    {
                        //db.DeleteOnTable("soalne", "loan_code='" + loan_code + "'");
                        //add_items(code);

                        success = true;
                    }
                    else
                    {
                        MessageBox.Show("Failed on saving.");
                    }
                }

                if (success)
                {
                    disp_list();
                    goto_tbcntrl_list();
                    tpg_info_enable(false);
                    frm_clear();
                }


            }
        }
        public void disp_list()
        {
            DataTable dt = null;
            dgv_list.Rows.Clear();

            dt = db.QueryBySQLCode("SELECT loan_code, loan_desc, loan_location, loan_type, user_id, whs_location_code, loan_cost_center_code,loan_cost_center_name,loan_sub_cost_center,employee_no,employee_name,loan_amount,loan_deduction,loan_transdate,deduction_date FROM rssys.hr_loanhdr ORDER BY loan_code");

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
                row.Cells["dgvl_deduction_date"].Value = dt.Rows[r]["deduction_date"].ToString();


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
            isnew = true;
            frm_clear();
            tpg_info_enable(true);
            goto_tbcntrl_info();
        }
        private void tpg_info_enable(Boolean flag)
        {
            txt_code.Enabled = flag;
        }
        private void goto_tbcntrl_info()
        {
            seltbp = true;
            tbcntrl_main.SelectedTab = tpg_info;
            tbcntrl_option.SelectedTab = tpg_opt_2;

            tpg_info.Show();
            tpg_opt_2.Show();
            seltbp = false;
        }

        public void frm_clear()
        {

            txt_code.Text = "";
            txt_desc.Text = "";
            cbo_stocklocation.SelectedIndex= -1;
            cbo_contraacct.SelectedIndex = -1;         
            cbo_stocklocation.SelectedIndex= -1;
            cbo_costcenter.SelectedIndex= -1;
            cbo_scc.SelectedIndex = -1;
            cbo_employee.SelectedIndex = -1;
            txt_amnt_loan.Text = "0.00";
            txt_deduction.Text = "0.00";
            dtp_trnxdt.Value = DateTime.Now;
            dtp_deduction.Value = DateTime.Now;
        }

        private void tbcntrl_main_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
                e.Cancel = true;
        }

        private void tbcntrl_option_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
                e.Cancel = true;
        }

        private void btn_upd_Click(object sender, EventArgs e)
        {
            int r = 0;
            String code = "";
            isnew = false;

            if (dgv_list.Rows.Count > 1)
            {
                try
                {
                   
                    r = dgv_list.CurrentRow.Index;
                    cbo_employee.SelectedValue = dgv_list["dgvl_empno", r].Value.ToString();
                    code = dgv_list["dgvl_code", r].Value.ToString();
                    txt_code.Text = code;
                    txt_desc.Text = dgv_list["dgvl_description", r].Value.ToString();
                    cbo_contraacct.SelectedValue = dgv_list["dgvl_trn_type", r].Value.ToString();
                    cbo_stocklocation.Text = dgv_list["dgvl_location", r].Value.ToString();
                    cbo_costcenter.Text = dgv_list["dgvl_ccname", r].Value.ToString();
                    cbo_scc.Text = dgv_list["dgvl_sub_costcenter", r].Value.ToString();
                    dtp_deduction.Value = gm.toDateValue(dgv_list["dgvl_deduction_date", r].Value.ToString());
                    dtp_trnxdt.Value = gm.toDateValue(dgv_list["dgvl_trnxdate", r].Value.ToString());
                    txt_amnt_loan.Text = gm.toNormalDoubleFormat(dgv_list["dgvl_loan_amount", r].Value.ToString()).ToString("0.00");
                    txt_deduction.Text = gm.toNormalDoubleFormat(dgv_list["dgvl_deduction_amount", r].Value.ToString()).ToString("0.00");
                  
                   
                }
                catch { }
               
                //disp_itemlist(code);
                //total_amountdue();
                goto_tbcntrl_info();

                //goto_win2();
            }
        }

        private void btn_mainexit_Click(object sender, EventArgs e)
        {
            goto_tbcntrl_list();
            tpg_info_enable(false);
            frm_clear();
        }

        private void dgv_list_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

       
    }
}
