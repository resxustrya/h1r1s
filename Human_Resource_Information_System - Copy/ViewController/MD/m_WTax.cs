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
    public partial class m_WTax : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();

        public m_WTax()
        {
            InitializeComponent();
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

        private void tbcntrl_option_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if(seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void btn_additem_Click(object sender, EventArgs e)
        {
            isnew = true;
            goto_win2();
            txt_code.Enabled = true;
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            goto_win1();
            CleanForm();
        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            Boolean success = false, ok = false;
            String notificationText = "";
            z_Notification notify = new z_Notification();

            String branch = GlobalClass.branch;
            String col = "", val = "";
            String notifyadd = null, id = "";
            String table = "hr_wtax";

            String code = "", description = "", exemption = "", bracket1 = "", bracket2 = "", bracket3 = "", bracket4 = "", bracket5 = "", bracket6 = "", bracket7 = "", bracket8 = "", bracket9 = "", bracket10 = "", factor1 = "", factor2 = "", factor3 = "", factor4 = "", factor5 = "", factor6 = "", factor7 = "", factor8 = "", factor9 = "", factor10 = "", add_on1 = "", add_on2 = "", add_on3 = "", add_on4 = "", add_on5 = "", add_on6 = "", add_on7 = "", add_on8 = "", add_on9 = "", add_on10 = "";

          
            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (String.IsNullOrEmpty(txt_description.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (String.IsNullOrEmpty(txt_exemption.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_bracket_1.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            if(txt_bracket_2.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            if(txt_bracket_3.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_bracket_4.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_bracket_5.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_bracket_6.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_bracket_7.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            if(txt_bracket_8.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_bracket_9.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_bracket_10.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            if(txt_factor_1.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_factor_2.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_factor_3.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_factor_4.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_factor_5.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_factor_6.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_factor_7.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_factor_8.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_factor_9.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_factor_10.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on1.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on2.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on3.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on4.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_add_on5.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_add_on6.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on7.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on8.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (txt_add_on9.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if(txt_add_on10.Text == "")
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            code = txt_code.Text;
            description = txt_description.Text;
            exemption = txt_exemption.Text;
            bracket1 = txt_bracket_1.Text;
            bracket2 = txt_bracket_2.Text;
            bracket3 = txt_bracket_3.Text;
            bracket4 = txt_bracket_4.Text;
            bracket5 = txt_bracket_5.Text;
            bracket6 = txt_bracket_6.Text;
            bracket7 = txt_bracket_7.Text;
            bracket8 = txt_bracket_8.Text;
            bracket9 = txt_bracket_9.Text;
            bracket10 = txt_bracket_10.Text;

            factor1 = txt_factor_1.Text;
            factor2 = txt_factor_2.Text;
            factor3 = txt_factor_3.Text;
            factor4 = txt_factor_4.Text;
            factor5 = txt_factor_5.Text;
            factor6 = txt_factor_6.Text;
            factor7 = txt_factor_7.Text;
            factor8 = txt_factor_8.Text;
            factor9 = txt_factor_9.Text;
            factor10 = txt_factor_10.Text;

            add_on1 = txt_add_on1.Text;
            add_on2 = txt_add_on2.Text;
            add_on3 = txt_add_on3.Text;
            add_on4 = txt_add_on4.Text;
            add_on5 = txt_add_on5.Text;
            add_on6 = txt_add_on6.Text;
            add_on7 = txt_add_on7.Text;
            add_on8 = txt_add_on8.Text;
            add_on9 = txt_add_on9.Text;
            add_on10 = txt_add_on10.Text;

            if(isnew)
            {
                col = "code,description,exemption,bracket1,bracket2,bracket3,bracket4,bracket5,bracket6,bracket7,bracket8,bracket9,bracket10,factor1,factor2,factor3,factor4,factor5,factor6,factor7,factor8,factor9,factor10,add_on1,add_on2,add_on3,add_on4,add_on5,add_on6,add_on7,add_on8,add_on9,add_on10";
                val = "" + db.str_E(code) + "," + db.str_E(description) + "," + db.str_E(exemption) + ",'" + bracket1 + "','" + bracket2 + "','" + bracket3 + "','" + bracket4 + "','" + bracket5 + "','" + bracket6 + "','" + bracket7 + "','" + bracket8 + "','" + bracket9 + "','" + bracket10 + "','" + factor1 + "','" + factor2 + "','" + factor3 + "','" + factor4 + "','" + factor5 + "','" + factor6 + "','" + factor7 + "','" + factor8 + "','" + factor9 + "','" + factor10 + "','" + add_on1 + "','" + add_on2 + "','" + add_on3 + "','" + add_on4 + "','" + add_on5 + "','" + add_on6 + "','" + add_on7 + "','" + add_on8 + "','" + add_on9 + "','" + add_on10 + "'";

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
                col = "description=" + db.str_E(description) + ",exemption=" + db.str_E(exemption) + ",bracket1='" + bracket1 + "',bracket2='" + bracket2 + "',bracket3='" + bracket3 + "',bracket4='" + bracket4 + "',bracket5='" + bracket5 + "',bracket6='" + bracket6 + "',bracket7='" + bracket7 + "',bracket8='" + bracket8 + "',bracket9='" + bracket9 + "',bracket10='" + bracket10 + "',factor1='" + factor1 + "',factor2='" + factor2 + "',factor3='" + factor3 + "',factor4='" + factor4 + "',factor5='" + factor5 + "',factor6='" + factor6 + "',factor7='" + factor7 + "',factor8='" + factor8 + "',factor9='" + factor9 + "',factor10='" + factor10 + "',add_on1='" + add_on1 + "',add_on2='" + add_on2 + "',add_on3='" + add_on3 + "',add_on4='" + add_on4 + "',add_on5='" + add_on5 + "',add_on6='" + add_on6 + "',add_on7='" + add_on7 + "',add_on8='" + add_on8 + "',add_on9='" + add_on9 + "',add_on10='" + add_on10 + "'";
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
                CleanForm();
                goto_win1();
                disp_list();

            }

        }
        private void disp_list()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;
            String type = "";
            DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_wtax WHERE COALESCE(cancel,cancel,'')<>'Y'");
            
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["code"].Value = dt.Rows[r]["code"].ToString();
                row.Cells["description"].Value = dt.Rows[r]["description"].ToString();
            }
        }

        private void m_WTax_Load(object sender, EventArgs e)
        {
            disp_list();
        }

        private void btn_upditem_Click(object sender, EventArgs e)
        {
            txt_code.Enabled = false;
            isnew = false;


            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String id = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(id))
                {

                    try{

                        DataTable dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_wtax WHERE code=" + db.str_E(id) + " AND COALESCE(cancel,cancel,'')<>'Y'");

                        txt_code.Text = dt.Rows[0]["code"].ToString();
                        txt_description.Text = dt.Rows[0]["description"].ToString();
                        txt_exemption.Text = dt.Rows[0]["exemption"].ToString();
                        txt_bracket_1.Text = dt.Rows[0]["bracket1"].ToString();
                        txt_bracket_2.Text = dt.Rows[0]["bracket2"].ToString();
                        txt_bracket_3.Text = dt.Rows[0]["bracket3"].ToString();
                        txt_bracket_4.Text = dt.Rows[0]["bracket4"].ToString();
                        txt_bracket_5.Text = dt.Rows[0]["bracket5"].ToString();
                        txt_bracket_6.Text = dt.Rows[0]["bracket6"].ToString();
                        txt_bracket_7.Text = dt.Rows[0]["bracket7"].ToString();
                        txt_bracket_8.Text = dt.Rows[0]["bracket8"].ToString();
                        txt_bracket_9.Text = dt.Rows[0]["bracket9"].ToString();
                        txt_bracket_10.Text = dt.Rows[0]["bracket10"].ToString();

                        txt_factor_1.Text = dt.Rows[0]["factor1"].ToString();
                        txt_factor_2.Text = dt.Rows[0]["factor2"].ToString();
                        txt_factor_3.Text = dt.Rows[0]["factor3"].ToString();
                        txt_factor_4.Text = dt.Rows[0]["factor4"].ToString();
                        txt_factor_5.Text = dt.Rows[0]["factor5"].ToString();
                        txt_factor_6.Text = dt.Rows[0]["factor6"].ToString();
                        txt_factor_7.Text = dt.Rows[0]["factor7"].ToString();
                        txt_factor_8.Text = dt.Rows[0]["factor8"].ToString();
                        txt_factor_9.Text = dt.Rows[0]["factor9"].ToString();
                        txt_factor_10.Text = dt.Rows[0]["factor10"].ToString();

                        txt_add_on1.Text = dt.Rows[0]["add_on1"].ToString();
                        txt_add_on2.Text = dt.Rows[0]["add_on2"].ToString();
                        txt_add_on3.Text = dt.Rows[0]["add_on3"].ToString();
                        txt_add_on4.Text = dt.Rows[0]["add_on4"].ToString();
                        txt_add_on5.Text = dt.Rows[0]["add_on5"].ToString();
                        txt_add_on6.Text = dt.Rows[0]["add_on6"].ToString();
                        txt_add_on7.Text = dt.Rows[0]["add_on7"].ToString();
                        txt_add_on8.Text = dt.Rows[0]["add_on8"].ToString();
                        txt_add_on9.Text = dt.Rows[0]["add_on9"].ToString();
                        txt_add_on10.Text = dt.Rows[0]["add_on10"].ToString();
   
                        
                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No witholding selected.");
                }
            }
            catch
            {
                MessageBox.Show("No witholding selected.");
            }

        }

        private void CleanForm()
        {
            txt_code.Text = "";
            txt_exemption.Text = "";
            txt_description.Text = "";


            txt_bracket_1.Text = "0.00";
            txt_bracket_2.Text = "0.00";
            txt_bracket_3.Text = "0.00";
            txt_bracket_4.Text = "0.00";
            txt_bracket_5.Text = "0.00";
            txt_bracket_6.Text = "0.00";
            txt_bracket_7.Text = "0.00";
            txt_bracket_8.Text = "0.00";
            txt_bracket_9.Text = "0.00";
            txt_bracket_10.Text = "0.00";

            txt_factor_1.Text = "0.00";
            txt_factor_2.Text = "0.00";
            txt_factor_3.Text = "0.00";
            txt_factor_4.Text = "0.00";
            txt_factor_5.Text = "0.00";
            txt_factor_6.Text = "0.00";
            txt_factor_7.Text = "0.00";
            txt_factor_8.Text = "0.00";
            txt_factor_9.Text = "0.00";
            txt_factor_10.Text = "0.00";

            txt_add_on1.Text = "0.00";
            txt_add_on2.Text = "0.00";
            txt_add_on3.Text = "0.00";
            txt_add_on4.Text = "0.00";
            txt_add_on5.Text = "0.00";
            txt_add_on6.Text = "0.00";
            txt_add_on7.Text = "0.00";
            txt_add_on8.Text = "0.00";
            txt_add_on9.Text = "0.00";
            txt_add_on10.Text = "0.00";
        }

        private void btn_delitem_Click(object sender, EventArgs e)
        {
            int r = -1;
            try
            {
                r = dgv_list.CurrentRow.Index;
                String wcode = dgv_list["code", r].Value.ToString();

                if (dgv_list.Rows.Count > 1 && !String.IsNullOrEmpty(wcode))
                {
                    if (MessageBox.Show("Are you sure you want to cancel this witholding tax?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {  
                        try
                        {
                            db.UpdateOnTable("hr_wtax", "cancel='Y'", "code=" + db.str_E(wcode) + "");

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
                    MessageBox.Show("No witholding selected.");
                }
            }
            catch
            {
                MessageBox.Show("No witholding selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {
            RPT_RES_entry frm = new RPT_RES_entry("MD07", "Witholding Tax Brackets");
            frm.print_master_data();
            frm.ShowDialog();
        }
    }
}
