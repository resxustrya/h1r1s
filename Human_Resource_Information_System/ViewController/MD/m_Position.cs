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
    public partial class m_Position : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db;
        public m_Position()
        {
            InitializeComponent();
        }

        private void m_Position_Load(object sender, EventArgs e)
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            db = new thisDatabase();
            disp_list();
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
            goto_win2();
            txt_code.Enabled = true;
            clear_frm();
        }


        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null;
            String table = "hr_position";


            String code = "", position = "";

            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (String.IsNullOrEmpty(txt_name.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            code = txt_code.Text;
            position = txt_name.Text;

            col = "postid, position_name";
            val = "" + db.str_E(code) + "," + db.str_E(position) + "";

            if (isnew)
            {

                db.DeleteOnTable(table, "postid=" + db.str_E(code) + " AND cancel='Y'");
                if (db.InsertOnTable(table, col, val))
                {
                    success = true;
                }
                else
                {
                    success = false;
                    //db.DeleteOnTable(table, "postid='" + code + "'");
                    MessageBox.Show("Failed on saving.");
                }
            }
            else
            {
                col = "postid=" + db.str_E(code) + ", position_name=" + db.str_E(position) + "";
                if (db.UpdateOnTable(table, col, "postid=" + db.str_E(code) + ""))
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
                clear_frm();
            }
        }
        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }

            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_position WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["ID"].Value = dt.Rows[r]["postid"].ToString();
                row.Cells["name"].Value = dt.Rows[r]["position_name"].ToString();

            }
        }
        private void clear_frm()
        {
            txt_code.Text = "";
            txt_name.Text = "";
        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            isnew = false;
            txt_code.Enabled = false;

            int r = -1;

            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["ID", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {
                    r = dgv_list.CurrentRow.Index;

                    try
                    {
                        txt_code.Text = id;
                        txt_name.Text = dgv_list["name", r].Value.ToString();
                    }
                    catch (Exception er) {  }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No position selected.");
                }

            }
            catch
            {
                MessageBox.Show("No position selected.");
            }




        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            String code = "";
            try
            {
                r = dgv_list.CurrentRow.Index;
                String postid = dgv_list["ID", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(postid))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this position?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_position", "cancel='Y'", "postid=" + db.str_E(postid) + "");

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
                    MessageBox.Show("No position selected.");
                }
            }
            catch
            {
                MessageBox.Show("No position selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD03", "Job Title List");
            frm.print_master_data();
            frm.ShowDialog();
        }
        
    }
}
