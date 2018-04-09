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
    public partial class t_addSched : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();
        t_BatchTimeLog frm_bt;
        t_manpowerschedule frm_wrk;
        public t_addSched()
        {
            InitializeComponent();
        }

        public t_addSched(t_BatchTimeLog bt)
        {
            InitializeComponent();
            frm_bt = bt;
        }
        public t_addSched(t_manpowerschedule wrk, Boolean _isnew)
        {
            InitializeComponent();
            frm_wrk = wrk;
            isnew = _isnew;
        }

        private void disp_info(){

            if (!isnew)
            {
                if (frm_wrk != null) {

                    int r = frm_wrk.getDGV().CurrentRow.Index;
                    DataGridViewRow row = frm_wrk.getDGV().Rows[r];

                    dtp_frm.Value = DateTime.Parse(row.Cells["dgvl_date"].Value.ToString());
                    dtp_to.Value = DateTime.Parse(row.Cells["dgvl_date"].Value.ToString());
                    dtp_timein1.Value = DateTime.Parse(row.Cells["dgvl_timein1"].Value.ToString());
                    dtp_timein2.Value = DateTime.Parse(row.Cells["dgvl_timein2"].Value.ToString());
                    dtp_timeout1.Value = DateTime.Parse(row.Cells["dgvl_timeout1"].Value.ToString());
                    dtp_timeout2.Value = DateTime.Parse(row.Cells["dgvl_timeout2"].Value.ToString());

                    cbo_shift_sched.SelectedValue = row.Cells["dgvl_shiftsched_code"].Value.ToString();

                }
            }
        }


        private void t_addSched_Load(object sender, EventArgs e)
        {
            gc.load_shift_schedule(cbo_shift_sched);

            dtp_timein1.CustomFormat = "hh:mm tt";
            dtp_timein2.CustomFormat = "hh:mm tt";
            dtp_timeout1.CustomFormat = "hh:mm tt";
            dtp_timeout2.CustomFormat = "hh:mm tt";

            disp_info();
        }
       

        private void btn_save_Click(object sender, EventArgs e)
        {

            //
            if (cbo_shift_sched.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a shift schedule.");
                cbo_shift_sched.DroppedDown = true;
                return;
            }

            String wrkd_frm = dtp_frm.Value.ToString("yyyy-MM-dd");
            String wrkd_to = dtp_to.Value.ToString("yyyy-MM-dd");
            String timein1 = dtp_timein1.Value.ToString("HH:mm"); 
            String timein2 = dtp_timein2.Value.ToString("HH:mm"); 
            String timeout1 = dtp_timeout1.Value.ToString("HH:mm"); 
            String timeout2 = dtp_timeout2.Value.ToString("HH:mm");
            String shift_sched = cbo_shift_sched.SelectedValue.ToString();
            String shift_sched_desc = cbo_shift_sched.Text;

            if (frm_wrk != null) {

                frm_wrk.set_work(isnew,wrkd_frm, wrkd_to, timein1, timein2, timeout1, timeout2, shift_sched, shift_sched_desc);

            }
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
