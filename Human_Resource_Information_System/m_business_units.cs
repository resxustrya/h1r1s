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
    public partial class m_business_units : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();


        public m_business_units()
        {
            InitializeComponent();
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
            if (seltbp == false)
            {
                e.Cancel = true;
            }
        }

        private void m_business_units_Load(object sender, EventArgs e)
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
            String table = "hr_business_unit";

            if (String.IsNullOrEmpty(txt_code.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }
            if (String.IsNullOrEmpty(txt_desc.Text))
            {
                MessageBox.Show("Please enter the required fields.");
                return;
            }

            String bucode = txt_code.Text;
            String bunit_desc = txt_desc.Text;
            String bank_disburse = txt_bank_disburse.Text;
            String bank_addr = txt_bank_addr.Text;
            String accnt_no = txt_accnt_no.Text;
            String contact_person = txt_contact_person.Text;
            String designation_cp = txt_designation_cp.Text;
            String bletter_prepared = txt_bletter_prepared.Text;
            String designation_blp = txt_designation_blp.Text;
            String bletter_noted = txt_bletter_noted.Text;
            String designation_bln = txt_designation_bln.Text;
            String accnt_data_folder = txt_accnt_data_folder.Text;
            String letter_format = cbo_letter_format.Text;
            
            if(isnew)
            {
                col = "bucode,bunit_desc,bank_disburse,bank_addr,accnt_no,contact_person,designation_cp, bletter_prepared ,designation_blp,bletter_noted,designation_bln,accnt_data_folder,letter_format";
                //val = "'" + code + "','" + bracket1 + "','" + bracket2 + "','" + s_credit + "','" + empshare_sc + "','" + s_ec + "','" + empshare_ec + "'";
                val = "" + db.str_E(bucode) + "," + db.str_E(bunit_desc) + "," + db.str_E(bank_disburse) + "," + db.str_E(bank_addr) + "," + db.str_E(accnt_no) + "," + db.str_E(contact_person) + "," + db.str_E(designation_cp) + ", " + db.str_E(bletter_prepared) + "," + db.str_E(designation_blp) + "," + db.str_E(bletter_noted) + "," + db.str_E(designation_bln) + "," + db.str_E(accnt_data_folder) + "," + db.str_E(letter_format) + "";


                db.DeleteOnTable(table, "bucode=" + db.str_E(bucode) + " AND cancel='Y'");
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
                col = "bunit_desc=" + db.str_E(bunit_desc) + ",bank_disburse=" + db.str_E(bank_disburse) + ",bank_addr=" + db.str_E(bank_addr) + ",accnt_no=" + db.str_E(accnt_no) + ",contact_person=" + db.str_E(contact_person) + ",designation_cp=" + db.str_E(designation_cp) + ", bletter_prepared=" + db.str_E(bletter_prepared) + ",designation_blp=" + db.str_E(designation_blp) + ",bletter_noted=" + db.str_E(bletter_noted) + ",designation_bln=" + db.str_E(designation_bln) + ",accnt_data_folder=" + db.str_E(accnt_data_folder) + ",letter_format=" + db.str_E(letter_format) + "";

                if (db.UpdateOnTable(table, col, "bucode=" + db.str_E(bucode) + ""))
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
            txt_desc.Text = "";
            txt_bank_disburse.Text = "";
            txt_bank_addr.Text = "";
            txt_accnt_no.Text = "";
            txt_contact_person.Text = "";
            txt_designation_cp.Text = "";
            txt_bletter_prepared.Text = "";
            txt_designation_blp.Text = "";
            txt_bletter_noted.Text = "";
            txt_designation_bln.Text = "";
            txt_accnt_data_folder.Text = "";
            cbo_letter_format.SelectedIndex = 0;
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

            DataTable dt = db.QueryBySQLCode("SELECT * from rssys.hr_business_unit WHERE COALESCE(cancel,cancel,'')<>'Y'");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                i = dgv_list.Rows.Add();
                DataGridViewRow row = dgv_list.Rows[i];

                row.Cells["dgvl_code"].Value = dt.Rows[r]["bucode"].ToString();
                row.Cells["dgvl_desc"].Value = dt.Rows[r]["bunit_desc"].ToString();
                row.Cells["dgvl_bank_disburse"].Value = dt.Rows[r]["bank_disburse"].ToString();
                row.Cells["dgvl_bank_addr"].Value = dt.Rows[r]["bank_addr"].ToString();
                row.Cells["dgvl_accnt_no"].Value = dt.Rows[r]["accnt_no"].ToString();
                row.Cells["dgvl_contact_person"].Value = dt.Rows[r]["contact_person"].ToString();
                row.Cells["dgvl_designation_cp"].Value = dt.Rows[r]["designation_cp"].ToString();
                row.Cells["dgvl_letter_format"].Value = dt.Rows[r]["letter_format"].ToString();
                row.Cells["dgvl_bletter_prepared"].Value = dt.Rows[r]["bletter_prepared"].ToString();
                row.Cells["dgvl_designation_blp"].Value = dt.Rows[r]["designation_blp"].ToString();
                row.Cells["dgvl_bletter_noted"].Value = dt.Rows[r]["bletter_noted"].ToString();
                row.Cells["dgvl_designation_bln"].Value = dt.Rows[r]["designation_bln"].ToString();
                row.Cells["dgvl_accnt_data_folder"].Value = dt.Rows[r]["accnt_data_folder"].ToString();

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
                        txt_desc.Text = dgv_list["dgvl_desc", r].Value.ToString();
                        txt_bank_disburse.Text = dgv_list["dgvl_bank_disburse", r].Value.ToString();
                        txt_bank_addr.Text = dgv_list["dgvl_bank_addr", r].Value.ToString();
                        txt_accnt_no.Text = dgv_list["dgvl_accnt_no", r].Value.ToString();
                        txt_contact_person.Text = dgv_list["dgvl_contact_person", r].Value.ToString();
                        txt_designation_cp.Text = dgv_list["dgvl_designation_cp", r].Value.ToString();
                        cbo_letter_format.Text = dgv_list["dgvl_letter_format", r].Value.ToString();
                        txt_bletter_prepared.Text = dgv_list["dgvl_bletter_prepared", r].Value.ToString();
                        txt_designation_blp.Text = dgv_list["dgvl_designation_blp", r].Value.ToString();
                        txt_bletter_noted.Text = dgv_list["dgvl_bletter_noted", r].Value.ToString();
                        txt_designation_bln.Text = dgv_list["dgvl_designation_bln", r].Value.ToString();
                        txt_accnt_data_folder.Text = dgv_list["dgvl_accnt_data_folder", r].Value.ToString();
	
                    }
                    catch (Exception er) { }

                    goto_win2();
                }
                else
                {
                    MessageBox.Show("No business unit selected.");
                }
            }
            catch
            {
                MessageBox.Show("No business unit selected.");
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
                    if (MessageBox.Show("Are you sure you want to cancel this business unit?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            db.UpdateOnTable("hr_business_unit", "cancel='Y'", "bucode=" + db.str_E(code) + "");

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
                    MessageBox.Show("No business unit selected.");
                }
            }
            catch
            {
                MessageBox.Show("No business unit selected.");
            }
        }

        private void btn_print_Click(object sender, EventArgs e)
        {

        }


    }
}
