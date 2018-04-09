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
    public partial class m_SSS : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();

        public m_SSS()
        {
            InitializeComponent();
        }
        

        private void tbcntrl_main_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if(seltbp == false)
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

        private void m_SSS_Load(object sender, EventArgs e)
        {
            
            disp_list();
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
            isnew = true;
            txt_code.Enabled = true;
            frm_clear();
            goto_win2();
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null, id = "";
            String table = "hr_sss";

            
            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            String code = txt_code.Text;
            String bracket1 = gm.toNormalDoubleFormat(txt_bracket1.Text).ToString("0.00");
            String bracket2 = gm.toNormalDoubleFormat(txt_bracket2.Text).ToString("0.00");  
            String s_credit = gm.toNormalDoubleFormat(txt_s_credit.Text).ToString("0.00"); 
            String empshare_sc = gm.toNormalDoubleFormat(txt_empshare_sc.Text).ToString("0.00");  
            String s_ec = gm.toNormalDoubleFormat(txt_s_ec.Text).ToString("0.00");  
            String empshare_ec = gm.toNormalDoubleFormat(txt_empshare_ec.Text).ToString("0.00");  

            if(isnew)
            {
                col = "code,bracket1,bracket2,s_credit,empshare_sc,s_ec,empshare_ec";
                val = "" + db.str_E(code) + ",'" + bracket1 + "','" + bracket2 + "','" + s_credit + "','" + empshare_sc + "','" + s_ec + "','" + empshare_ec + "'";

                db.DeleteOnTable(table, "code=" + db.str_E(code) + " AND cancel='Y'");
                if (db.InsertOnTable(table, col, val))
                {
                    success = true;
                }
                else
                {
                    success = false;
                    //db.DeleteOnTable(table, "code='" + code + "'");
                    MessageBox.Show("Failed on saving.");
                }
            }
            else
            {
                col = "bracket1='" + bracket1 + "',bracket2='" + bracket2 + "',s_credit='" + s_credit + "',empshare_sc='" + empshare_sc + "',s_ec='" + s_ec + "',empshare_ec='" + empshare_ec + "'";

                if (db.UpdateOnTable(table, col, "code=" + db.str_E(code) + ""))
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
                frm_clear();
                disp_list();
            }
        }

        private void frm_clear()
        {
            txt_code.Text = "";
            txt_bracket1.Text = "0.00";
            txt_bracket2.Text = "0.00";
            txt_s_credit.Text = "0.00";
            txt_empshare_ec.Text = "0.00";
            txt_s_ec.Text = "0.00";
            txt_empshare_sc.Text = "0.00";
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
        }


        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;

            DataTable dt = db.QueryBySQLCode("SELECT * from rssys.hr_sss WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["code"].ToString();
                row.Cells["dgvl_bracket1"].Value = gm.toAccountingFormat(dt.Rows[r]["bracket1"].ToString());
                row.Cells["dgvl_bracket2"].Value = gm.toAccountingFormat(dt.Rows[r]["bracket2"].ToString()); 
                row.Cells["dgvl_s_credit"].Value = gm.toAccountingFormat(dt.Rows[r]["s_credit"].ToString()); 
                row.Cells["dgvl_empshare_sc"].Value = gm.toAccountingFormat(dt.Rows[r]["empshare_sc"].ToString()); 
                row.Cells["dgvl_s_ec"].Value = gm.toAccountingFormat(dt.Rows[r]["s_ec"].ToString()); 
                row.Cells["dgvl_empshare_ec"].Value = gm.toAccountingFormat(dt.Rows[r]["empshare_ec"].ToString()); 
               
            }
        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            isnew = false;
            txt_code.Enabled = false;

            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["dgvl_code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {

                    try
                    {
                        txt_code.Text = id;
                        txt_bracket1.Text = dgv_list["dgvl_bracket1", r].Value.ToString();
                        txt_bracket2.Text = dgv_list["dgvl_bracket2", r].Value.ToString();
                        txt_s_credit.Text = dgv_list["dgvl_s_credit", r].Value.ToString();
                        txt_empshare_ec.Text = dgv_list["dgvl_empshare_sc", r].Value.ToString();
                        txt_s_ec.Text = dgv_list["dgvl_s_ec", r].Value.ToString();
                        txt_empshare_sc.Text = dgv_list["dgvl_empshare_ec", r].Value.ToString();

                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No SSS selected.");
                }
            }
            catch
            {
                MessageBox.Show("No SSS selected.");
            }


        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {

            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String code = dgv_list["dgvl_code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(code))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this SSS?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_sss", "cancel='Y'", "code=" + db.str_E(code) + "");

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
                    MessageBox.Show("No SSS selected.");
                }
            }
            catch
            {
                MessageBox.Show("No SSS selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD08", "SSS Brackets");
            frm.print_master_data();
            frm.ShowDialog();
        }

        private void txt_number_AccountingFormat_onKeyUp(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back || e.KeyChar == ',' || e.KeyChar == '.')))
            {
                e.Handled = true;
            }
            if(e.KeyChar == '.'){
                if(((TextBox)sender).Text.Contains('.')){
                    e.Handled = true;
                }
            }
        }


    }
}
