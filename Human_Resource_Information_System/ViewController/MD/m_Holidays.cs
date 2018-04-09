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
    public partial class m_Holidays : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db;
        public m_Holidays()
        {
            InitializeComponent();
        }


        private void m_Holidays_Load(object sender, EventArgs e)
        {
            db = new thisDatabase();

            load_hol_type();
            disp_list();
        
        }
        private void load_hol_type()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("code");
            dt.Columns.Add("type");
            dt.Rows.Add();
            dt.Rows[0]["code"] = "L";
            dt.Rows[0]["type"] = "Legal";
            dt.Rows.Add();
            dt.Rows[1]["code"] = "S";
            dt.Rows[1]["type"] = "Special";

            cbo_holiday_type.DataSource = dt;
            cbo_holiday_type.DisplayMember = "type";
            cbo_holiday_type.ValueMember = "code";
            cbo_holiday_type.SelectedIndex = -1;
        }

        private void tbcntrl_option_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if(seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void tbcntrl_main_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void btn_additem_Click(object sender, EventArgs e)
        {
            txt_id.Text = "";
            goto_win2();
            isnew = true;
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
            String notifyadd = null,id="";
            String table = "hr_holidays";


            String date = "", description = "", type = "" ;

            if (txt_description.Text == "")
            {
                MessageBox.Show("Please enter holiday description.");
                return;
            }
            if(cbo_holiday_type.SelectedIndex == -1)
            {
                MessageBox.Show("Plase select holiday type.");
                cbo_holiday_type.DroppedDown = true;
                return;
            }
            if(txt_id.Text != "")
            {
                id = txt_id.Text;
            }
            date = dtp_holiday.Value.ToShortDateString();
            description = txt_description.Text;
            type = cbo_holiday_type.SelectedValue.ToString();
           
            if (isnew)
            {
                col = "date_holiday,description,holiday_type";
                val = "'" + date + "'," + db.str_E(description) + ",'" + type + "'";


                db.DeleteOnTable(table, "id=" + db.str_E(id) + " AND cancel='Y'");
                if (db.InsertOnTable(table, col, val))
                {
                    success = true;
                }
                else
                {
                    success = false;
                    //db.DeleteOnTable(table, "id='" + id + "'");
                    MessageBox.Show("Failed on saving.");
                }
            }
            else
            {
                col = "date_holiday='" + date + "', description=" + db.str_E(description) + ",holiday_type='" + type + "'";

                if (db.UpdateOnTable(table, col, "id=" + db.str_E(id) + ""))
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
                disp_list();
                clear_frm();
            }
        }

        private void tpg_info_Click(object sender, EventArgs e)
        {

        }
        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            String type = "";
            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_holidays WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["date"].Value = DateTime.Parse(dt.Rows[r]["date_holiday"].ToString()).ToString("yyyy-MM-dd");
                row.Cells["description"].Value = dt.Rows[r]["description"].ToString();
                if (dt.Rows[r]["holiday_type"].ToString() == "L")
                {
                    row.Cells["type"].Value = "LEGAL";
                }
                else if (dt.Rows[r]["holiday_type"].ToString() == "S")
                {
                    row.Cells["type"].Value = "SPECIAL";
                }
                row.Cells["code"].Value = dt.Rows[r]["holiday_type"].ToString();
                row.Cells["id"].Value = dt.Rows[r]["id"].ToString();

            }
        }
        private void clear_frm()
        {
            dtp_holiday.ResetText();
            txt_description.Text = "";
            cbo_holiday_type.SelectedIndex = -1;
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
            clear_frm();
        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            isnew = false;
            
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["id", r].Value.ToString(); 

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {

                    try
                    {
                        txt_id.Text = id;
                        dtp_holiday.Value = Convert.ToDateTime(dgv_list["date", r].Value.ToString());
                        txt_description.Text = dgv_list["description", r].Value.ToString();
                        cbo_holiday_type.SelectedValue = dgv_list["code", r].Value.ToString();
                    }
                    catch (Exception er) {  }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No holiday selected.");
                }
            }
            catch
            {
                MessageBox.Show("No holiday selected.");
            }
        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String holid = dgv_list["id", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(holid))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this holiday?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_holidays", "cancel='Y'", "id='" + holid + "'");

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
                    MessageBox.Show("No holiday selected.");
                }
            }
            catch
            {
                MessageBox.Show("No holiday selected.");
            }
        }

        private void txt_name_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD05", "Holidays");
            frm.print_master_data();
            frm.ShowDialog();
        }
    }
}
