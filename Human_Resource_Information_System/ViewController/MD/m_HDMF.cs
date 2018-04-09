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
    public partial class m_HDMF : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();
        public m_HDMF()
        {
            InitializeComponent();
        }

        private void btn_additem_Click(object sender, EventArgs e)
        {
            isnew = true;
            txt_code.Enabled = true;
            goto_win2();
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

        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;

            DataTable dt = db.QueryBySQLCode("SELECT * from rssys.hr_hdmf WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];
             
                row.Cells["code"].Value = dt.Rows[r]["code"].ToString();
                row.Cells["bracket1"].Value = dt.Rows[r]["bracket1"].ToString();
                row.Cells["bracket2"].Value = dt.Rows[r]["bracket2"].ToString();
                row.Cells["pct"].Value = dt.Rows[r]["pct"].ToString();
                row.Cells["emp_ee"].Value = dt.Rows[r]["emp_ee"].ToString();
                row.Cells["emp_er"].Value = dt.Rows[r]["emp_er"].ToString();
               

            }
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null, id = "";
            String table = "hr_hdmf";

            String code = "", bracket1 = "", bracket2 = "", pct = "", emp_ee = "", emp_er = "";

            if(String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            code = txt_code.Text;
            bracket1 = txt_bracket1.Text;
            bracket2 = txt_bracket2.Text;
            pct = gm.toNormalDoubleFormat(txt_pct.Text).ToString("0.00");
            emp_ee = gm.toNormalDoubleFormat(txt_emp_ee.Text).ToString("0.00");
            emp_er = gm.toNormalDoubleFormat(txt_emp_er.Text).ToString("0.00");

            if (isnew)
            {
                col = "code,bracket1,bracket2,pct,emp_ee,emp_er";
                val = "" + db.str_E(code) + ",'" + bracket1 + "','" + bracket2 + "','" + emp_ee + "','" + pct + "','" + emp_er + "'";

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
                col = "bracket1='" + bracket1 + "',bracket2='" + bracket2 + "',pct='" + pct + "',emp_er='" + emp_er + "',emp_ee='" + emp_ee + "'";
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
            if(success)
            {
                goto_win1();
                disp_list();
                ClearForm();
            }
        }
        private void ClearForm()
        {
            txt_code.Text = "";
            txt_bracket1.Text = "0.00";
            txt_bracket2.Text = "0.00";
            txt_pct.Text = "0.00";
        }

        private void m_HDMF_Load(object sender, EventArgs e)
        {
            disp_list();
        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            isnew = false;
            txt_code.Enabled = false;


            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {
                    try
                    {

                        txt_code.Text = id;
                        txt_bracket1.Text = dgv_list["bracket1", r].Value.ToString();
                        txt_bracket2.Text = dgv_list["bracket2", r].Value.ToString();
                        txt_pct.Text = dgv_list["pct", r].Value.ToString();
                        txt_emp_ee.Text = dgv_list["emp_ee", r].Value.ToString();
                        txt_emp_er.Text = dgv_list["emp_er", r].Value.ToString();

                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No HDMF selected.");
                }
            }
            catch
            {
                MessageBox.Show("No HDMF selected.");
            }

        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
            ClearForm();
        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String hcode = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(hcode))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this HDMF?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_hdmf", "cancel='Y'", "code=" + db.str_E(hcode) + "");
                                 
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
                    MessageBox.Show("No HDMF selected.");
                }
            }
            catch
            {
                MessageBox.Show("No HDMF selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD10", "Pag-ibig Brackets");
            frm.print_master_data();
            frm.ShowDialog();
        }

    }
}
