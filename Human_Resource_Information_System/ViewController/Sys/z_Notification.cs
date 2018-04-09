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
    public partial class z_Notification : Form
    {
        thisDatabase db;
        GlobalClass gc;
        String thisNotification = "";
        String thisFromWhatForm = "";

        public z_Notification()
        {
            InitializeComponent();

            db = new thisDatabase();
            gc = new GlobalClass();

            gc.load_branch(cbo_branch);

            this.Load += z_Notification_Load;
            dtp_frm.ValueChanged += dateTimePicker1_ValueChanged;
            dtp_to.ValueChanged += dateTimePicker2_ValueChanged;
        }

        void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            toDisplay();
        }

        void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            toDisplay();
        }

        void btn_new_Click(object sender, EventArgs e)
        {
            toDisplay();
        }


        private void z_Notification_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            toDisplay();
        }

        private void toDisplay()
        {
            DateTime currentDate = dtp_frm.Value.Date;
            DateTime nextDate = dtp_to.Value.Date;
            int r = 0;
            DataTable dt = db.get_notification(currentDate, nextDate);

            try
            {
                dgv_list.Rows.Clear();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    r = dgv_list.Rows.Add();
                    DataGridViewRow row = dgv_list.Rows[r];

                    row.Cells["t_date"].Value = DateTime.Parse(dt.Rows[i]["t_date"].ToString()).ToString("yyyy-MM-dd");
                    row.Cells["t_time"].Value = dt.Rows[i]["t_time"].ToString();
                    row.Cells["username"].Value = dt.Rows[i]["username"].ToString();
                    row.Cells["module"].Value = dt.Rows[i]["module"].ToString();
                    row.Cells["notification_text"].Value = dt.Rows[i]["notification_text"].ToString();
                }
            }
            catch
            {

            }
        }

        public void saveNotification(String notification, String form)
        {
            Boolean success = false;
            this.thisNotification = notification;
            this.thisFromWhatForm = form;

            try
            {
                success = db.InsertOnTable("notification", "notification_text, username, module, t_date, t_time", "'" + thisNotification + "', '" + GlobalClass.username + "', '" + thisFromWhatForm + "', '" + db.get_systemdate("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm") + "'");
            }
            catch { }
        }

        private void dtp_frm_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dtp_to_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
