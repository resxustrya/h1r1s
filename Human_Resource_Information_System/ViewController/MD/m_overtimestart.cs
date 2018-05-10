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
    public partial class m_overtimestart : Form
    {
        thisDatabase db = new thisDatabase();
        public m_overtimestart()
        {
            InitializeComponent();
        }

        private void m_overtimestart_Load(object sender, EventArgs e)
        {
            DataTable dt = db.QueryBySQLCode("Select * FROM rssys.hr_ot_start LIMIT 1");
            time_start.Value = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + dt.Rows[0]["time_start"].ToString());
            time_start.CustomFormat = "hh:mm tt";
            time_start.ShowUpDown = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            db.QueryBySQLCode("DELETE FROM rssys.hr_ot_start");
            String code = "", time = "", col = "", val = "", table = "hr_ot_start";
            code = db.get_pk("ot_id");
            time = time_start.Value.ToString("HH:mm");
            col = "ot_id,time_start";
            val = "'" + code + "','" + time + "'";
            if (db.InsertOnTable(table, col, val))
            {
                db.set_pkm99("ot_id", db.get_nextincrementlimitchar(code, 8)); //changes from 'hr_empid'
                MessageBox.Show("Successfully Updated");
                this.Close();
            }
            else
            {
                MessageBox.Show("Failed on saving.");
            }
        }
    }
}
