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
      
    public partial class m_payrollperiod : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();

        public m_payrollperiod()
        {
            InitializeComponent();
        }

        private void tpg_info_Click(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void m_payrollperiod_Load(object sender, EventArgs e)
        {
            load_pay_classic();

            disp_list();
        }

        private void frm_clear()
        {

            try
            {
                cbo_d_phil.SelectedIndex = 1;
                cbo_dw_tax.SelectedIndex = 1;
                cbo_d_sss_c.SelectedIndex = 1;
                cbo_d_phil.SelectedIndex = 1;
                cbo_pagibig.SelectedIndex = 1;
                cbo_p_type.SelectedIndex = 1;
                cbo_payroll_classic.SelectedIndex = 1;
            }
            catch { }
            dtp_from.ResetText();
            dtp_to.ResetText();

            txt_code.Text = "";
            txt_f_year.Text = DateTime.Now.ToString("yyyy");
            txt_f_month.Text = DateTime.Now.ToString("MM");
            txt_num_days.Text = "";

        }


        private void tbcntrl_main_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void tbcntrl_option_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void goto_win2()
        {
            seltbp = true;
            tbcntrl_option.SelectedTab = tpg_opt_2;
            tpg_opt_2.Show();

            tbcntrl_main.SelectedTab = tpg_info;
            tpg_info.Show();
            seltbp = false;
        }

        private void goto_win1()
        {
            seltbp = true;
            tbcntrl_option.SelectedTab = tpg_opt_1;
            tpg_opt_1.Show();

            tbcntrl_main.SelectedTab = tpg_list;
            tpg_list.Show();
            seltbp = false;
        }

        private void btn_additem_Click(object sender, EventArgs e)
        {
            txt_code.Enabled = true;
            isnew = true;
            frm_clear();
            goto_win2();
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
        }

      

        private void load_pay_classic()
        {
            cbo_payroll_classic.DataSource = null;
            try
            {
                DataTable dt = new DataTable();


                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_payroll_classic WHERE COALESCE(cancel,cancel,'')<>'Y'");
                if (dt.Rows.Count > 0)
                {
                    cbo_payroll_classic.DataSource = dt;
                    cbo_payroll_classic.DisplayMember = "description";
                    cbo_payroll_classic.ValueMember = "code";
                    cbo_payroll_classic.SelectedIndex = -1;
                }
            }
            catch (Exception) { }
        }

        private void btn_save_Click_1(object sender, EventArgs e)
        {
           
            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null, id = "";
            String table = "hr_payrollpariod";

            String pay_code = "", date_from = "", date_to = "", d_w_tax = "", d_sss_c = "", d_philhealth = "", d_pagibig = "", financial_year = "", month = "", pay_type = "", num_days = "", payroll_classic="";


            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(String.IsNullOrEmpty(txt_f_year.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (String.IsNullOrEmpty(txt_f_month.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(String.IsNullOrEmpty(txt_num_days.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

         
            if (cbo_dw_tax.SelectedIndex == 0)
            {
                d_w_tax = "Y";
            }
            else
            {
                d_w_tax = "N";
            }
            if (cbo_d_sss_c.SelectedIndex == 0)
            {
                d_sss_c = "Y";
            }
            else
            {
                d_sss_c = "N";
            }
            if (cbo_d_phil.SelectedIndex == 0)
            {
                d_philhealth = "Y";
            }
            else
            {
                d_philhealth = "N";
            }
            if (cbo_pagibig.SelectedIndex == 0)
            {
                d_pagibig = "Y";
            }
            else
            {
                d_pagibig = "N";
            }
            if (cbo_p_type.SelectedIndex == 0)
            {
                pay_type = "S";
            }
            else
            {
                pay_type = "R";
            }

            pay_code = txt_code.Text;
            date_from = dtp_from.Value.ToShortDateString();
            date_to = dtp_to.Value.ToShortDateString();
            financial_year = txt_f_year.Text;
            num_days = txt_num_days.Text;
            month = txt_f_month.Text;
            if (String.IsNullOrEmpty(month))
            {
                month = DateTime.Now.ToString("MM");    
            }
            if (cbo_payroll_classic.SelectedIndex != -1)
            {
                payroll_classic = cbo_payroll_classic.SelectedValue.ToString();
            }
           

            if (isnew)
            {

                 col = "pay_code,date_from,date_to,d_w_tax,d_sss_c ,d_philhealth,d_pagibig,financial_year,month,pay_type,num_days,payroll_classic";
                val = "" + db.str_E(pay_code) + ",'" + date_from + "','" + date_to + "','" + d_w_tax + "','" + d_sss_c + "','" + d_philhealth + "','" + d_pagibig + "','" + financial_year + "','" + month + "','" + pay_type + "','" + num_days + "','" + payroll_classic + "'";

                db.DeleteOnTable(table, "pay_code=" + db.str_E(pay_code) + " AND cancel='Y'");
                if (db.InsertOnTable(table, col, val))
                {
                    success = true;
                }
                else
                {
                    success = false;
                   //db.DeleteOnTable(table, "pay_code='" + pay_code +"'"); 
                    MessageBox.Show("Failed on saving.");
                }
            }
            else
            {
                col = "date_from='" + date_from + "',date_to='" + date_to + "',d_w_tax='" + d_w_tax + "',d_sss_c='" + d_sss_c + "',d_philhealth='" + d_philhealth + "',d_pagibig='" + d_pagibig +  "',financial_year='" + financial_year + "',month='" + month + "',pay_type='" + pay_type + "',num_days='" + num_days + "',payroll_classic='" + payroll_classic + "'"; 
                if (db.UpdateOnTable(table, col, "pay_code=" + db.str_E(pay_code) + ""))
                {
                    success = true;
                }
                else
                {
                    success = false;
                    MessageBox.Show("Failed on saving.");
                }
            }

            if (success)
            {
                goto_win1();
                disp_list();
               
            }

        }


        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;
            String type = "";
            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_payrollpariod WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["code"].Value = dt.Rows[r]["pay_code"].ToString();
                row.Cells["date_from"].Value = DateTime.Parse(dt.Rows[r]["date_from"].ToString()).ToString("yyyy-MM-dd");
                row.Cells["date_to"].Value = DateTime.Parse(dt.Rows[r]["date_to"].ToString()).ToString("yyyy-MM-dd"); 
            }
        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            txt_code.Enabled = false;
            isnew = false;

            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {
                    try
                    {
                        DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_payrollpariod WHERE pay_code=" + db.str_E(id) + " AND COALESCE(cancel,cancel,'')<>'Y'");

                        txt_code.Text = dt.Rows[0]["pay_code"].ToString();
                        dtp_from.Value = Convert.ToDateTime(dt.Rows[0]["date_from"].ToString());
                        dtp_to.Value = Convert.ToDateTime(dt.Rows[0]["date_to"].ToString());
                        if (dt.Rows[0]["d_w_tax"].ToString() == "Y")
                        {
                            cbo_dw_tax.SelectedIndex = 0;
                        }
                        else
                        {
                            cbo_dw_tax.SelectedIndex = 1;
                        }
                        if (dt.Rows[0]["d_sss_c"].ToString() == "Y")
                        {
                            cbo_d_sss_c.SelectedIndex = 0;
                        }
                        else
                        {
                            cbo_d_sss_c.SelectedIndex = 1;
                        }
                        if (dt.Rows[0]["d_philhealth"].ToString() == "Y")
                        {
                            cbo_d_phil.SelectedIndex = 0;
                        }
                        else
                        {
                            cbo_d_phil.SelectedIndex = 1;
                        }
                        if (dt.Rows[0]["d_pagibig"].ToString() == "Y")
                        {
                            cbo_pagibig.SelectedIndex = 0;
                        }
                        else
                        {
                            cbo_pagibig.SelectedIndex = 1;
                        }
                        txt_f_year.Text = dt.Rows[0]["financial_year"].ToString();
                        txt_f_month.Text = dt.Rows[0]["month"].ToString();

                        if (dt.Rows[0]["pay_type"].ToString() == "S")
                        {
                            cbo_p_type.SelectedIndex = 0;
                        }
                        else
                        {
                            cbo_p_type.SelectedIndex = 1;
                        }
                        txt_num_days.Text = dt.Rows[0]["num_days"].ToString();
                        cbo_payroll_classic.SelectedValue = dt.Rows[0]["payroll_classic"].ToString();
                    }
                    catch (Exception er) {  }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No payroll selected.");
                }
            }
            catch
            {
                MessageBox.Show("No payroll selected.");
            }

        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String pay_code = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(pay_code))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this payroll?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_payrollpariod", "cancel='Y'", "pay_code=" + db.str_E(pay_code) + "");

                            disp_list();
                            MessageBox.Show("Cancelled successfully");
                        }
                        catch
                        {
                            MessageBox.Show("Invalid to cancel.");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No payroll selected.");
                }
            }
            catch
            {
                MessageBox.Show("No payroll selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD06", "Payroll Period");
            frm.print_master_data();
            frm.ShowDialog(); 
        }

        private void txt_f_month_TextChanged(object sender, EventArgs e)
        {
            if(txt_f_month.Text.Length != 0)
            {
                try
                {
                    double n = gm.toNormalDoubleFormat(txt_f_month.Text);
                    if (0 < n && n <= 12)
                    {
                        txt_f_month.Text = n.ToString("0");
                        return;
                    }
                } catch { }
                txt_f_month.Text = DateTime.Now.ToString("MM");
            }
        }
    }
}
