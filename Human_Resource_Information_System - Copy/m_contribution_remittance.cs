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
    public partial class m_contribution_remittance : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();

        public m_contribution_remittance()
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

        private void m_contribution_remittance_Load(object sender, EventArgs e)
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
            String table = "hr_contri_remittance";

            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (String.IsNullOrEmpty(txt_month.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            String code = txt_code.Text;
            String month = txt_code.Text;
            String sbr = txt_sbr.Text;
            String pbr = txt_pbr.Text;
            String pr = txt_pr.Text;

            String sbr_date = dtp_sbr_date.Value.ToString("yyyy-MM-dd");
            String pbr_date = dtp_pbr_date.Value.ToString("yyyy-MM-dd");
            String pr_date = dtp_pr_date.Value.ToString("yyyy-MM-dd");
            
            if(isnew)
            {
                col = "crcode,month,sbr,sbr_date,pbr,pbr_date,pr,pr_date";
                val = ""+db.str_E(code)+","+db.str_E(month)+","+db.str_E(sbr)+",'"+sbr_date+"',"+db.str_E(pbr)+",'"+pbr_date+"',"+db.str_E(pr)+",'"+pr_date+"'";

                db.DeleteOnTable(table, "crcode=" + db.str_E(code) + " AND cancel='Y'");
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
                col = "crcode=" + db.str_E(code) + ",month=" + db.str_E(month) + ",sbr=" + db.str_E(sbr) + ",pbr=" + db.str_E(pbr) + ",pr=" + db.str_E(pr) + ",sbr_date='" + sbr_date + "',pbr_date='" + pbr_date + "',pr_date='" + pr_date + "'";

                if (db.UpdateOnTable(table, col, "crcode=" + db.str_E(code) + ""))
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
            txt_month.Text = "";
            txt_sbr.Text = "";
            txt_pbr.Text = "";
            txt_pr.Text = "";

            dtp_sbr_date.ResetText();
            dtp_pbr_date.ResetText();
            dtp_pr_date.ResetText();
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

            DataTable dt = db.QueryBySQLCode("SELECT * from rssys.hr_contri_remittance WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["crcode"].ToString();
                row.Cells["dgvl_month"].Value = dt.Rows[r]["month"].ToString();
                row.Cells["dgvl_sbr"].Value = dt.Rows[r]["sbr"].ToString();
                row.Cells["dgvl_sbr_date"].Value = gm.toDateString(dt.Rows[r]["sbr_date"].ToString(), "");
                row.Cells["dgvl_pbr"].Value = dt.Rows[r]["pbr"].ToString();
                row.Cells["dgvl_pbr_date"].Value = gm.toDateString(dt.Rows[r]["pbr_date"].ToString(), "");
                row.Cells["dgvl_pr"].Value = dt.Rows[r]["pr"].ToString();
                row.Cells["dgvl_pr_date"].Value = gm.toDateString(dt.Rows[r]["pr_date"].ToString(), ""); 
               
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
                String code = dgv_list["dgvl_code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(code))
                {

                    try
                    {
                        txt_code.Text = code;
                        txt_month.Text = dgv_list["dgvl_month", r].Value.ToString();
                        txt_sbr.Text = dgv_list["dgvl_sbr", r].Value.ToString();
                        txt_pbr.Text = dgv_list["dgvl_pbr", r].Value.ToString();
                        txt_pr.Text = dgv_list["dgvl_pr", r].Value.ToString();

                        dtp_sbr_date.Value = gm.toDateValue(dgv_list["dgvl_sbr_date", r].Value.ToString());
                        dtp_pbr_date.Value = gm.toDateValue(dgv_list["dgvl_pbr_date", r].Value.ToString());
                        dtp_pr_date.Value = gm.toDateValue(dgv_list["dgvl_pr_date", r].Value.ToString());
                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No remittance unit selected.");
                }
            }
            catch
            {
                MessageBox.Show("No remittance unit selected.");
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
                    if (MessageBox.Show("Are you sure you want to cancel this contribution remittance?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_contri_remittance", "cancel='Y'", "crcode=" + db.str_E(code) + "");

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
                    MessageBox.Show("No remittance selected.");
                }
            }
            catch
            {
                MessageBox.Show("No remittance selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {

        }


    }
}
