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
    public partial class m_Section : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db;
        public m_Section()
        {
            InitializeComponent();
        }

        private void m_Section_Load(object sender, EventArgs e)
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            db = new thisDatabase();

            gc.load_dept(cbo_department);

            disp_list();
        }

        private void tbcntrl_main_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
                e.Cancel = true;
        }

        private void btn_additem_Click(object sender, EventArgs e)
        {
            isnew = true;
            goto_win2();
            clear_frm();
            txt_code.Enabled = true;
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

        private void tbcntrl_option_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
                e.Cancel = true;
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            Boolean success = false,ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null;
            String table = "hr_depsection";
          

            String code = "", name = "",deptid = "";

            if(String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(String.IsNullOrEmpty(txt_name.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(cbo_department.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a department.");
                cbo_department.DroppedDown = true;
                return;
            }

            code = txt_code.Text;
            name = txt_name.Text;
            deptid = cbo_department.SelectedValue.ToString();

            col = "secid,section_name,deptid";
            val = "" + db.str_E(code) + "," + db.str_E(name) + ",'" + deptid + "'";

            if (isnew)
            {

                db.DeleteOnTable(table, "secid=" + db.str_E(code) + " AND cancel='Y'");
                if (db.InsertOnTable(table, col, val))
                {
                    success = true;
                }
                else
                {
                    success = false;
                    //db.DeleteOnTable(table, "secid='" + code + "'");
                    MessageBox.Show("Failed on saving.");
                }
            }
            else
            {
                col = "secid=" + db.str_E(code) + ", section_name=" +db.str_E( name) + ",deptid='" + deptid + "'";

                if (db.UpdateOnTable(table, col, "secid=" + db.str_E(code) + ""))
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
        private void clear_frm()
        {
            txt_code.Text = "";
            txt_name.Text = "";
        }

        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }

            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_depsection WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["ID"].Value = dt.Rows[r]["secid"].ToString();
                row.Cells["name"].Value = dt.Rows[r]["section_name"].ToString();
                row.Cells["deptid"].Value = dt.Rows[r]["deptid"].ToString();

                cbo_department.SelectedValue = row.Cells["deptid"].Value;
                row.Cells["department"].Value = cbo_department.Text;
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
                String id = dgv_list["ID", r].Value.ToString();

                if (dgv_list.Rows.Count > 1)
                {
                    try
                    {
                        txt_code.Text = id;
                        txt_name.Text = dgv_list["name", r].Value.ToString();
                        cbo_department.SelectedValue = dgv_list["deptid", r].Value.ToString();
                    }
                    catch (Exception er) {}

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No section selected.");
                }
            }
            catch
            {
                MessageBox.Show("No section selected.");
            }
        }


        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String secid = dgv_list["ID", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(secid))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this section?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try {
                            db.UpdateOnTable("hr_depsection", "cancel='Y'", "secid=" + db.str_E(secid) + "");

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
                    MessageBox.Show("No section selected.");
                }
            }
            catch
            {
                MessageBox.Show("No section selected.");
            }
        
        }


        private void txt_code_TextChanged(object sender, EventArgs e)
        {

        }

        private void dgv_list_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD02", "Department Section List");
            frm.print_master_data();
            frm.ShowDialog();
        }
    }
}
