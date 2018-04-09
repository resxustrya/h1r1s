using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Human_Resource_Information_System
{
    public partial class activation_form : Form
    {
        public Login frm;
        string mac_addr = "";
        string cipher = "", decipher = "";
        thisDatabase db = new thisDatabase();
        public activation_form()
        {
            InitializeComponent();
        }
        public activation_form(Login _frm, String mac_addr, String decipher)
        {
            InitializeComponent();
            this.frm = _frm;
            this.mac_addr = mac_addr;
            this.decipher = decipher;
        }
        private void btn_activate_Click(object sender, EventArgs e)
        {
            String licensed_pc = "", col = "", val = "", key = "", pk = "";
            if (txt_key.Text == string.Empty)
            {
                MessageBox.Show("Please Enter License key.");
                txt_key.Focus();
            }
            else
            {

                DataTable dt, dt2 = new DataTable();
                dt = db.QueryBySQLCode("SELECT * FROM rssys.x09 WHERE license_id='" + txt_key.Text + "'");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["licensed_pc"].ToString() != string.Empty && dt.Rows[0]["license_id"].ToString() != string.Empty)
                    {
                        MessageBox.Show("Activation Failed, Key is Already Used.", "Activation Warning",
    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        col = "licensed_pc='" + decipher + "'";
                        if (db.UpdateOnTable("x09", col, "license_id='" + dt.Rows[0]["license_id"].ToString() + "'"))
                        {
                            MessageBox.Show("Activation Success.", "Activation Complete.",
    MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Activation Failed, Something went wrong.", "Activation Error",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            frm.ok = "ok";
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Activation Failed, Key is Invalid.", "Activation Error",
       MessageBoxButtons.OK, MessageBoxIcon.Error);

                }


            }

            this.Close();
        }

        private void activation_form_Load(object sender, EventArgs e)
        {
            txt_key.Focus();
        }

        private void btn_activate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                String licensed_pc = "", col = "", val = "", key = "", pk = "";
                if (txt_key.Text == string.Empty)
                {
                    MessageBox.Show("Please Enter License key.");
                    txt_key.Focus();
                }
                else
                {

                    DataTable dt, dt2 = new DataTable();
                    dt = db.QueryBySQLCode("SELECT * FROM rssys.x09 WHERE license_id='" + txt_key.Text + "'");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["licensed_pc"].ToString() != string.Empty && dt.Rows[0]["license_id"].ToString() != string.Empty)
                        {
                            MessageBox.Show("Activation Failed, Key is Already Used.", "Activation Warning",
        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            col = "licensed_pc='" + decipher + "'";
                            if (db.UpdateOnTable("x09", col, "license_id='" + dt.Rows[0]["license_id"].ToString() + "'"))
                            {
                                MessageBox.Show("Activation Success.", "Activation Complete.",
        MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Activation Failed, Something went wrong.", "Activation Error",
        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                frm.ok = "ok";
                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("Activation Failed, Key is Invalid.", "Activation Error",
           MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }


                }

                this.Close();
            }
        }
    }
}
