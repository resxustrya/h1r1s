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
    public partial class m_PhilHealth : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();
        public m_PhilHealth()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

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

        private void m_PhilHealth_Load(object sender, EventArgs e)
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
            String table = "hr_philhealth";

            String code = "", bracket1 = "", bracket2 = "", salary_base = "", emp_er = "", emp_ee = "";

            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            code = txt_code.Text;
            bracket1 = txt_bracket1.Text;
            bracket2 = txt_bracket2.Text;
            salary_base = txt_salary_base.Text;
            emp_er = txt_emp_er.Text;
            emp_ee = txt_emp_ee.Text;

            if(isnew)
            {
                col = "code,bracket1,bracket2,salary_base,emp_er,emp_ee";
                val = "" + db.str_E(code) + ",'" + bracket1 + "','" + bracket2 + "','" + salary_base + "','" + emp_er + "','" + emp_ee + "'";

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
                col = "bracket1='" + bracket1 + "',bracket2='" + bracket2 + "',salary_base='" + salary_base + "',emp_er='" + emp_er + "',emp_ee='" + emp_ee + "'";

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
            txt_salary_base.Text = "0.00";
            txt_emp_ee.Text = "0.00";
            txt_emp_er.Text = "0.00";
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

            DataTable dt = db.QueryBySQLCode("SELECT * from rssys.hr_philhealth WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];
                
                row.Cells["code"].Value = dt.Rows[r]["code"].ToString();
                row.Cells["bracket1"].Value = dt.Rows[r]["bracket1"].ToString();
                row.Cells["bracket2"].Value = dt.Rows[r]["bracket2"].ToString();
                row.Cells["salary_base"].Value = dt.Rows[r]["salary_base"].ToString();
                row.Cells["emp_er"].Value = dt.Rows[r]["emp_er"].ToString();
                row.Cells["emp_ee"].Value = dt.Rows[r]["emp_ee"].ToString();
                i++;
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
                String id = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {

                    try
                    {
                        txt_code.Text = id;
                        txt_bracket1.Text = dgv_list["bracket1", r].Value.ToString();
                        txt_bracket2.Text = dgv_list["bracket2", r].Value.ToString();
                        txt_salary_base.Text = dgv_list["salary_base", r].Value.ToString();
                        txt_emp_ee.Text = dgv_list["emp_ee", r].Value.ToString();
                        txt_emp_er.Text = dgv_list["emp_er", r].Value.ToString();

                    }
                    catch (Exception er) { }

                    goto_win2();

                }
                else
                {
                    MessageBox.Show("No philhealth selected.");
                }
            }
            catch
            {
                MessageBox.Show("No philhealth selected.");
            }


        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {

            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String phcode = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(phcode))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this philhealth?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_philhealth", "cancel='Y'", "code=" + db.str_E(phcode) + "");

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
                    MessageBox.Show("No philhealth selected.");
                }
            }
            catch
            {
                MessageBox.Show("No philhealth selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD09", "PhilHealth Brackets");
            frm.print_master_data();
            frm.ShowDialog();
        }


    }
}
