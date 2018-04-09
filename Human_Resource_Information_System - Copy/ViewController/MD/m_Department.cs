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
    public partial class m_Department : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db;
        public m_Department()
        {
            InitializeComponent();
        }

        private void m_Department_Load(object sender, EventArgs e)
        {
            gc = new GlobalClass();
            gm = new GlobalMethod();
            db = new thisDatabase();
            
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
            String table = "hr_department";


            String code = "", name = "";

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
            name = txt_name.Text;

            col = "deptid,dept_name";
            val = "" + db.str_E(code) + "," + db.str_E(name) + "";

            if (isnew)
            {

                db.DeleteOnTable(table, "deptid='" + code + "' AND cancel='Y'");// use to replace new data in cancel PK
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
                col = "deptid=" + db.str_E(code) + ", dept_name=" + db.str_E(name) + "";

                if (db.UpdateOnTable(table, col, "deptid=" + db.str_E(code) + ""))
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

            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_department WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["ID"].Value = dt.Rows[r]["deptid"].ToString();
                row.Cells["name"].Value = dt.Rows[r]["dept_name"].ToString();

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
                String deptid = dgv_list["ID", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(deptid))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this department?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        
                        try
                        {
                            db.UpdateOnTable("hr_department", "cancel='Y'", "deptid=" + db.str_E(deptid) + "");

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
                    MessageBox.Show("No department selected.");
                }
            }
            catch 
            {
                MessageBox.Show("No department selected.");
            }

        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            isnew = false;
            txt_code.Enabled = false;

            int r = -1;
            String code = "";
            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["ID", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {

                    try
                    {
                        txt_code.Text = id;
                        txt_name.Text = dgv_list["name", r].Value.ToString();

                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No department selected.");
                }
            }
            catch { MessageBox.Show("No department selected.");  }

           
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD01", "Department List");
            frm.print_master_data();
            frm.ShowDialog();
        }


    }
}
