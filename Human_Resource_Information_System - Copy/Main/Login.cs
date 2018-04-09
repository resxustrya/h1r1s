using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Deployment.Application;

namespace Human_Resource_Information_System
{
    public partial class Login : Form
    {
        thisDatabase db = new thisDatabase();

        public Login()
        {
            InitializeComponent();

            try
            {
                db = new thisDatabase();

                lbl_server.Text = thisDatabase.servers;
            }
            catch (Exception er)
            {
                MessageBox.Show("Wrong Database information. \n" + er.Message);
            }
        }

        private void btn_access_Click(object sender, EventArgs e)
        {
            enter_login();
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            try
            {
                String brn = db.get_m99branch();

                cbo_branch.SelectedIndex = Convert.ToInt32(brn) - 1;

                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    lbl_version.Text = string.Format("Version {0}", ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4));
                }
            }
            catch { }
        }

        private void txt_user_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                enter_login();
            }
        }

        private void txt_pass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                enter_login();
            }
        }

        private void enter_login()
        {
            String branch = db.get_m99branch();
            String comp = db.get_m99comp_name();
            String user = db.validate_login(txt_user.Text, txt_pass.Text, comp);

            if (String.IsNullOrEmpty(user) == false)
            {
                GlobalClass.username = user;
                GlobalClass.branch = branch;
                GlobalClass.server_ip = lbl_server.Text;
                DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Invalid username and password. Pls try again.");
                txt_pass.Text = "";
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //System.Diagnostics.Process.Start("https://web.facebook.com/RightAppsSolutions/");

            MessageBox.Show("For information, contact 0915-806-0792 / 0942-734-7599.\nAlso like us on facebook \n https://web.facebook.com/RightAppsOfficial/");
        }
    }
}
