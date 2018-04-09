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
    public partial class m_jobtitle : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();

        public m_jobtitle()
        {
            InitializeComponent();
        }

        private void m_jobtitle_Load(object sender, EventArgs e)
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



        private void btn_additem_Click(object sender, EventArgs e)
        {
            goto_win2();
            isnew = true;
            frm_clear();
            txt_code.Enabled = true;
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
            String table = "hr_jobtitle";


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

            String code = txt_code.Text;
            String name = txt_name.Text;

            col = "jtid,jtitle_name";
            val = "" + db.str_E(code) + "," + db.str_E(name) + "";

            if (isnew)
            {

                db.DeleteOnTable(table, "jtid=" + db.str_E(code) + " AND cancel='Y'");// use to replace new data in cancel PK
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
                col = "jtid=" + db.str_E(code) + ", jtitle_name=" + db.str_E(name) + "";

                if (db.UpdateOnTable(table, col, "jtid=" + db.str_E(code) + ""))
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

            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_jobtitle WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];
                // jtid jtitle_name
                row.Cells["dgvl_code"].Value = dt.Rows[r]["jtid"].ToString();
                row.Cells["dgvl_name"].Value = dt.Rows[r]["jtitle_name"].ToString();

            }
        }

        private void frm_clear()
        {
            txt_code.Text = "";
            txt_name.Text = "";
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
                    if (MessageBox.Show("Are you sure you want to cancel this job title?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        
                        try
                        {
                            db.UpdateOnTable("hr_jobtitle", "cancel='Y'", "jtid=" + db.str_E(code) + "");

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
                    MessageBox.Show("No job title selected.");
                }
            }
            catch 
            {
                MessageBox.Show("No job title selected.");
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
                        txt_name.Text = dgv_list["dgvl_name", r].Value.ToString();

                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No job title selected.");
                }
            }
            catch { MessageBox.Show("No job title selected."); }

           
        }

        private void btn_print_Click(object sender, UICuesEventArgs e)
        {

        }

        private void btn_print_Click(object sender, EventArgs e)
        {

        }


    }
}
