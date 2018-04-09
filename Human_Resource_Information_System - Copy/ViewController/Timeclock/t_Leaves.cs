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
    public partial class t_Leaves : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();

        public t_Leaves()
        {
            InitializeComponent();
        }

        private void t_Leaves_Load(object sender, EventArgs e)
        {
            gc.load_leave_type(cbo_leave);
            gc.load_employee(cbo_employee);

            disp_list();
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
            if(seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void btn_additem_Click(object sender, EventArgs e)
        {
            txt_code.Enabled = true;
            isnew = true;
            goto_win2();
            frm_clear();
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

        private void btn_save_Click(object sender, EventArgs e)
        {

            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null, id = "";
            String table = "hr_leaves";

            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an employee.");
                cbo_employee.DroppedDown = true;
                return;
            }
            if (cbo_leave_pay.Text == "YES" && gm.toNormalDoubleFormat(txt_amount.Text) == 0)
            {
                MessageBox.Show("Please input leave amount.");
                return;
            }

            String lvcode = txt_code.Text;
            String no_of_days = Double.Parse(txt_no_of_days.Text).ToString();

            String d_filed = dtp_filed.Value.ToString("yyyy-MM-dd");
            String leave_from = dtp_lfrm.Value.ToString("yyyy-MM-dd");
            String leave_to = dtp_lto.Value.ToString("yyyy-MM-dd");
            String leave_amount = txt_amount.Text;
            String empid = (cbo_employee.SelectedValue ?? "").ToString();
            String leave_pay = cbo_leave_pay.Text;
            String leave_type = (cbo_leave.SelectedValue ?? "").ToString();

            String frm_am = Convert.ToString(chk_fam.Checked);
            String frm_pm = Convert.ToString(chk_fpm.Checked);
            String to_am = Convert.ToString(chk_tam.Checked);
            String to_pm = Convert.ToString(chk_tpm.Checked);

            // hr_leaves, lvcode, empid, d_filed, leave_from, leave_to, frm_am, frm_pm, to_am, to_pm, no_of_days, leave_pay, leave_type
            if (isnew)
            {
                col = "lvcode, empid, d_filed, leave_from, leave_to, frm_am, frm_pm, to_am, to_pm, no_of_days, leave_pay, leave_type, leave_amount";
                val = "'" + lvcode + "', '" + empid + "', '" + d_filed + "', '" + leave_from + "', '" + leave_to + "', '" + frm_am + "', '" + frm_pm + "', '" + to_am + "', '" + to_pm + "', '" + no_of_days + "', '" + leave_pay + "', '" + leave_type + "', '" + leave_amount + "'";

                db.DeleteOnTable(table, "lvcode='" + lvcode + "' AND cancel='Y'");
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
                col = "empid='" + empid + "', d_filed='" + d_filed + "', leave_from='" + leave_from + "', leave_to='" + leave_to + "', frm_am='" + frm_am + "', frm_pm='" + frm_pm + "', to_am='" + to_am + "', to_pm='" + to_pm + "', no_of_days='" + no_of_days + "', leave_pay='" + leave_pay + "', leave_type='" + leave_type + "', leave_amount='" + leave_amount + "'";
                if (db.UpdateOnTable(table, col, "lvcode='" + lvcode + "'"))
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
            txt_no_of_days.Text = "0";

            chk_fam.Checked = true;
            chk_fpm.Checked = true;
            chk_tam.Checked = true;
            chk_tpm.Checked = true;

            dtp_filed.ResetText();
            dtp_lfrm.ResetText();
            dtp_lto.ResetText();

            cbo_leave_pay.Text = "NO";
            txt_amount.Text = "0.00";
            cbo_employee.SelectedIndex = -1;

            try{ cbo_leave.SelectedIndex = 0; }
            catch { cbo_leave.SelectedIndex = -1; }
        }

        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;

            DataTable dt = db.QueryBySQLCode("SELECT * from rssys.hr_leaves WHERE COALESCE(cancel,cancel,'')<>'Y' ORDER BY d_filed DESC");
            
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["lvcode"].ToString();
                row.Cells["dgvl_empid"].Value = dt.Rows[r]["empid"].ToString();
                row.Cells["dgvl_leave_type_code"].Value = dt.Rows[r]["leave_type"].ToString();

                cbo_employee.SelectedValue = row.Cells["dgvl_empid"].Value;
                row.Cells["dgvl_employee"].Value = cbo_employee.Text;
                cbo_leave.SelectedValue = row.Cells["dgvl_leave_type_code"].Value;
                row.Cells["dgvl_leave_type"].Value = cbo_leave.Text;

                row.Cells["dgvl_leave_pay"].Value = dt.Rows[r]["leave_pay"].ToString();

                row.Cells["dgvl_d_filed"].Value = gm.toDateString(dt.Rows[r]["d_filed"].ToString(),"");
                row.Cells["dgvl_leave_from"].Value = gm.toDateString(dt.Rows[r]["leave_from"].ToString(), "");
                row.Cells["dgvl_lfrm_am"].Value = Convert.ToBoolean(dt.Rows[r]["frm_am"].ToString()) ? "✔" : "";
                row.Cells["dgvl_lfrm_pm"].Value = Convert.ToBoolean(dt.Rows[r]["frm_pm"].ToString()) ? "✔" : "";
                row.Cells["dgvl_leave_to"].Value = gm.toDateString(dt.Rows[r]["leave_to"].ToString(),"");
                row.Cells["dgvl_lto_am"].Value = Convert.ToBoolean(dt.Rows[r]["to_am"].ToString()) ? "✔" : "";
                row.Cells["dgvl_lto_pm"].Value = Convert.ToBoolean(dt.Rows[r]["to_pm"].ToString()) ? "✔" : "";


                row.Cells["dgvl_amount"].Value = gm.toNormalDoubleFormat(dt.Rows[r]["leave_amount"].ToString()).ToString("0.00");

                row.Cells["dgvl_no_of_days"].Value = dt.Rows[r]["no_of_days"].ToString();

                //✔
                i++;
            }
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
            frm_clear();
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

                        txt_no_of_days.Text = dgv_list["dgvl_no_of_days", r].Value.ToString();

                        chk_fam.Checked = !String.IsNullOrEmpty(dgv_list["dgvl_lfrm_am", r].Value.ToString());
                        chk_fpm.Checked = !String.IsNullOrEmpty(dgv_list["dgvl_lfrm_pm", r].Value.ToString());
                        chk_tam.Checked = !String.IsNullOrEmpty(dgv_list["dgvl_lto_am", r].Value.ToString());
                        chk_tpm.Checked = !String.IsNullOrEmpty(dgv_list["dgvl_lto_pm", r].Value.ToString());

                        dtp_filed.Value = DateTime.Parse(dgv_list["dgvl_d_filed", r].Value.ToString());
                        dtp_lfrm.Value = DateTime.Parse(dgv_list["dgvl_leave_from", r].Value.ToString()); 
                        dtp_lto.Value = DateTime.Parse(dgv_list["dgvl_leave_to", r].Value.ToString());

                        cbo_leave_pay.Text = dgv_list["dgvl_leave_pay", r].Value.ToString();
                        cbo_employee.SelectedValue = dgv_list["dgvl_empid", r].Value.ToString();
                        cbo_leave.SelectedValue = dgv_list["dgvl_leave_type_code", r].Value.ToString();

                        txt_amount.Text = gm.toNormalDoubleFormat(dgv_list["dgvl_amount", r].Value.ToString()).ToString("0.00");
                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No leave selected.");
                }
            }
            catch
            {
                MessageBox.Show("No leave selected.");
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
                    if (MessageBox.Show("Are you sure you want to cancel this leave?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_leaves", "cancel='Y'", "lvcode='" + code + "'");

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
                    MessageBox.Show("No leave selected.");
                }
            }
            catch
            {
                MessageBox.Show("No leave selected.");
            }
        }

        private void cbo_leave_pay_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_amount.Enabled = cbo_leave_pay.Text == "YES";
            if (!txt_amount.Enabled)
            {
                txt_amount.Text = "0.00";
            }
        }


    }
}
