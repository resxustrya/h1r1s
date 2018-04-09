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
    public partial class m_ShiftSchedule : Form
    {

        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db;
        public m_ShiftSchedule()
        {
            InitializeComponent();
        }

        private void m_ShiftSchedule_Load(object sender, EventArgs e)
        {
            db = new thisDatabase();
            gc = new GlobalClass();
            gm = new GlobalMethod();

            dtp_am_in.CustomFormat = "hh:mm tt";
            dtp_am_in.ShowUpDown = true;

            dtp_am_out.CustomFormat = "hh:mm tt";
            dtp_am_out.ShowUpDown = true;

            dtp_pm_in.CustomFormat = "hh:mm tt";
            dtp_pm_in.ShowUpDown = true;

            dtp_pm_out.CustomFormat = "hh:mm tt";
            dtp_pm_out.ShowUpDown = true;

            disp_list();
        }


        private void btn_additem_Click(object sender, EventArgs e)
        {
            txt_code.Enabled = true;
            isnew = true;
            frm_clear();
            goto_win2();
        }

        private void tbcntrl_option_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if(seltbp == false)
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

        private void tbcntrl_main_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if(seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
        }

        private void frm_clear()
        {
            txt_code.Text = "";
            txt_remark.Text = "";

            dtp_am_in.Value = DateTime.Now;
            dtp_am_out.Value = DateTime.Now;
            dtp_pm_in.Value = DateTime.Now;
            dtp_pm_out.Value = DateTime.Now;
            dtp_pm_in.Enabled = false;
            dtp_pm_out.Enabled = false;

            chk_am.Checked = true;
            chk_pm.Checked = false;
            cbo_type.SelectedIndex = 0;
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null;
            String table = "hr_shift_schedule";

            String code = "", am_timein = "", am_timeout = "", pm_timein = "", pm_timeout = "", remarks = "", type = "";

            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (!chk_am.Checked && !chk_pm.Checked)
            {
                MessageBox.Show("Please enter time in/time out.");
                return;
            }
            if (cbo_type.SelectedIndex == -1)
            {
                MessageBox.Show("Please select type.");
                cbo_type.DroppedDown = true;
                return;
            }

            if (chk_am.Checked)
            {
                am_timein = dtp_am_in.Value.ToString("HH:mm");
                am_timeout = dtp_am_out.Value.ToString("HH:mm");
            }
            if (chk_pm.Checked)
            {
                pm_timein = dtp_pm_in.Value.ToString("HH:mm");
                pm_timeout = dtp_pm_out.Value.ToString("HH:mm");
            }


            code = txt_code.Text;
            remarks = txt_remark.Text;
            type = cbo_type.Text;

            col = "code,am_timein,am_timeout,pm_timein,pm_timeout,remarks,type";
            val = "" + db.str_E(code) + ",'" + am_timein + "','" + am_timeout + "','" + pm_timein + "','" + pm_timeout + "','" + remarks + "','" + type + "'";

            if (isnew)
            {

                db.DeleteOnTable(table, "code=" + db.str_E(code) + " AND cancel='Y'");// use to replace new data in cancel PK
                if (db.InsertOnTable(table, col, val))
                {
                    success = true;
                }
                else
                {
                    success = false;
                    //db.DeleteOnTable(table, "deptid='" + code + "'");
                    MessageBox.Show("Failed on saving.");
                }

            }
            else
            {
                col = "code=" + db.str_E(code) + ",am_timein='" + am_timein + "',am_timeout='" + am_timeout + "',pm_timein='" + pm_timein + "',pm_timeout='" + pm_timeout + "',remarks='" + remarks + "',type='" + type + "'";

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
                disp_list();
                goto_win1();
                frm_clear();
            }

        }

        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }

            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_shift_schedule WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["code"].Value = dt.Rows[r]["code"].ToString();
                row.Cells["timein_am"].Value = dt.Rows[r]["am_timein"].ToString();
                row.Cells["timeout_am"].Value = dt.Rows[r]["am_timeout"].ToString();
                row.Cells["timein_pm"].Value = dt.Rows[r]["pm_timein"].ToString();
                row.Cells["timeout_pm"].Value = dt.Rows[r]["pm_timeout"].ToString();
                row.Cells["remarks"].Value = dt.Rows[r]["remarks"].ToString();
                row.Cells["type"].Value = dt.Rows[r]["type"].ToString();

            }
        }


        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String code = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(code))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this shift schedule?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        try
                        {
                            db.UpdateOnTable("hr_shift_schedule", "cancel='Y'", "code='" + code + "'");

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
                    MessageBox.Show("No shift schedule selected.");
                }
            }
            catch
            {
                MessageBox.Show("No shift schedule selected.");
            }

        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            txt_code.Enabled = false;
            isnew = false;
            frm_clear();


            try
            {
                int r = dgv_list.CurrentRow.Index;
                String code = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(code))
                {
                    try
                    {
                        txt_code.Text = code;
                        if (!String.IsNullOrEmpty(dgv_list["timein_am", r].Value.ToString()) && !String.IsNullOrEmpty(dgv_list["timeout_am", r].Value.ToString()))
                        {
                            dtp_am_in.Value = DateTime.Parse(dgv_list["timein_am", r].Value.ToString());
                            dtp_am_out.Value = DateTime.Parse(dgv_list["timeout_am", r].Value.ToString());
                            change_chk_am(true);
                        }

                        if (!String.IsNullOrEmpty(dgv_list["timein_pm", r].Value.ToString()) && !String.IsNullOrEmpty(dgv_list["timeout_pm", r].Value.ToString()))
                        {
                            dtp_pm_in.Value = DateTime.Parse(dgv_list["timein_pm", r].Value.ToString());
                            dtp_pm_out.Value = DateTime.Parse(dgv_list["timeout_pm", r].Value.ToString());
                            change_chk_pm(true);
                        }
                        txt_remark.Text = dgv_list["remarks", r].Value.ToString();
                        cbo_type.Text = dgv_list["type", r].Value.ToString();
                    }
                    catch { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No shift schedule selected.");
                }
            }
            catch
            {
                MessageBox.Show("No shift schedule selected.");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void chk_pm_CheckedChanged(object sender, EventArgs e)
        {
            change_chk_pm(chk_pm.Checked);
        }
        private void change_chk_pm(Boolean chk)
        {
            chk_pm.Checked = chk;
            dtp_pm_in.Enabled = chk;
            dtp_pm_out.Enabled = chk;
        }

        private void chk_am_CheckedChanged(object sender, EventArgs e)
        {
            change_chk_am(chk_am.Checked);
        }
        private void change_chk_am(Boolean chk)
        {
            chk_am.Checked = chk;
            dtp_am_in.Enabled = chk_am.Checked;
            dtp_am_out.Enabled = chk_am.Checked;
        }
        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD04", "Shift Schedule List");
            frm.print_master_data();
            frm.ShowDialog();
        }


    }
}
