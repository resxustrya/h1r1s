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
    public partial class RPT_RES_entry2 : Form
    {
        //Report rpt;
        thisDatabase db = new thisDatabase();
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        String rpt_code = "";

        Boolean isReady = false;
        public RPT_RES_entry2(String _rpt_code)
        {
            InitializeComponent();
            rpt_code = _rpt_code;

            if (rpt_code == "")
            {
                this.Text = "Employee Register";
            
            
            }
            else if (rpt_code == "P101" || rpt_code == "P102")
            {
                if (rpt_code == "P101") this.Text = "Payroll Summary Report";
                else this.Text = "Payslip";

                lbl_1.Text = "Payroll Period";
                lbl_2.Text = "Department";
                lbl_3.Text = "Specific Employee";
                lbl_4.Text = "Report Type";
                lbl_5.Text = "Sort Option";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show(); lbl_4.Show(); lbl_5.Show();

                chk_4.Text = "Hold employee only";
                chk_5.Text = "Detailed report";
                chk_4.Show(); chk_5.Show();

                load_payroll_period(cbo_1);
                load_department(cbo_2frm); load_department(cbo_2to);
                load_employee(cbo_3);
                cbo_1.Show(); cbo_2frm.Show(); cbo_2to.Show(); cbo_3.Show();

                cbo_4.Items.Add("All employees");
                cbo_4.Items.Add("Monthly paid");
                cbo_4.Items.Add("Daily paid");
                cbo_4.SelectedIndex = 0;
                cbo_5.Items.Add("By department/section");
                cbo_5.Items.Add("By payment type");
                cbo_5.Items.Add("By department/employee");
                cbo_5.SelectedIndex = 0;
                cbo_4.Show(); cbo_5.Show();

            }


            else if (rpt_code == "P202")
            {
                this.Text = "SSS Contribution Summary";

                lbl_1.Text = "Financial Year";
                lbl_2.Text = "Month";
                lbl_3.Text = "Department From";
                lbl_4.Text = "           To";
                lbl_1.Show(); lbl_2.Show(); lbl_2to.Show(); lbl_3.Show(); lbl_4.Show();

                cbo_1.DropDownStyle = ComboBoxStyle.Simple;
                cbo_2frm.DropDownStyle = ComboBoxStyle.Simple;
                cbo_2to.DropDownStyle = ComboBoxStyle.Simple;
                cbo_1.Text = DateTime.Now.ToString("yyyy");
                cbo_2frm.Text = "1";
                cbo_2to.Text = DateTime.Now.ToString("MM");
                cbo_1.Show(); cbo_2frm.Show(); cbo_2to.Show();

                load_department(cbo_3); load_department(cbo_4l);
                cbo_3.Show(); cbo_4l.Show();
            }
            else if (rpt_code == "P203")
            {
                this.Text = "PhilHealth Contribution Summary";
                
                lbl_1.Text = "Financial Year";
                lbl_2.Text = "Month";
                lbl_3.Text = "Department From";
                lbl_4.Text = "           To";
                lbl_1.Show(); lbl_2.Show(); lbl_2to.Show(); lbl_3.Show(); lbl_4.Show();

                cbo_1.DropDownStyle = ComboBoxStyle.Simple;
                cbo_2frm.DropDownStyle = ComboBoxStyle.Simple;
                cbo_2to.DropDownStyle = ComboBoxStyle.Simple;
                cbo_1.Text = DateTime.Now.ToString("yyyy");
                cbo_2frm.Text = "1";
                cbo_2to.Text =  DateTime.Now.ToString("MM");
                cbo_1.Show(); cbo_2frm.Show(); cbo_2to.Show();

                load_department(cbo_3); load_department(cbo_4l);
                cbo_3.Show(); cbo_4l.Show();
            }
            else if (rpt_code == "E103")
            {
                this.Text = "Payslip";

                lbl_1.Text = "Department From";
                lbl_2.Text = "           To";
                lbl_1.Show(); lbl_2.Show(); 

                load_department(cbo_1); load_department(cbo_2);
                cbo_1.Show(); cbo_2.Show();

                chk_3.Text = "Include employee rate history";
                chk_41.Text = "Confidential Only";
                chk_3.Show(); chk_41.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }

            else if (rpt_code == "M102")
            {
                this.Text = "Other Earning Summary";

                lbl_1.Text = "Payroll Period";
                lbl_2.Text = "Other Earning From";
                lbl_3.Text = "              To";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show();

                load_payroll_period(cbo_1);
                load_other_income(cbo_2);
                load_other_income(cbo_3);
                cbo_1.Show(); cbo_2.Show(); cbo_3.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }
            else if (rpt_code == "M103")
            {
                this.Text = "Other Deduction Summary";

                lbl_1.Text = "Payroll Period";
                lbl_2.Text = "Other Deduction From";
                lbl_3.Text = "                To";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show();

                load_payroll_period(cbo_1);
                load_other_deduction(cbo_2);
                load_other_deduction(cbo_3);
                cbo_1.Show(); cbo_2.Show(); cbo_3.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }
            else if (rpt_code == "M104")
            {
                this.Text = "Other Deduction per Employee";

                lbl_1.Text = "Payroll Period From";
                lbl_2.Text = "               To";
                lbl_3.Text = "Other Deduction";
                lbl_4.Text = "Employee";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show(); lbl_4.Show();

                load_payroll_period(cbo_1);
                load_payroll_period(cbo_2);
                load_other_deduction(cbo_3);
                load_employee(cbo_4);
                cbo_1.Show(); cbo_2.Show(); cbo_3.Show(); cbo_4.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }

            else if (rpt_code == "T101")
            {
                this.Text = "Employee DTR Report";

                lbl_1.Text = "Department From";
                lbl_2.Text = "             To";
                lbl_3.Text = "Work Dates";
                lbl_4.Text = "Specific Employee";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show(); lbl_3to.Show(); lbl_4.Show();

                load_department(cbo_1);
                load_department(cbo_2);
                load_employee(cbo_4l);
                cbo_1.Show(); cbo_2.Show(); cbo_4l.Show();

                dtp_3frm.Value = DateTime.Parse(dtp_3frm.Value.ToString("yyyy-MM-01"));
                dtp_3frm.Show(); dtp_3to.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }
            else if (rpt_code == "T102")
            {
                this.Text = "Daily Timelog Report";

                lbl_1.Text = "Department From";
                lbl_2.Text = "             To";
                lbl_3.Text = "Work Dates";
                lbl_4.Text = "Specific Employee";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show(); lbl_3to.Show(); lbl_4.Show();

                load_department(cbo_1);
                load_department(cbo_2);
                load_employee(cbo_4l);
                cbo_1.Show(); cbo_2.Show(); cbo_4l.Show();

                dtp_3frm.Value = DateTime.Parse(dtp_3frm.Value.ToString("yyyy-MM-01"));
                dtp_3frm.Show(); dtp_3to.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }

            else if (rpt_code == "T103")
            {
                this.Text = "Absences, Late, and Undertime";

                lbl_1.Text = "Department From";
                lbl_2.Text = "             To";
                lbl_3.Text = "Work Dates";
                lbl_4.Text = "Specific Employee";
                lbl_1.Show(); lbl_2.Show(); lbl_3.Show(); lbl_3to.Show(); lbl_4.Show();

                load_department(cbo_1);
                load_department(cbo_2);
                load_employee(cbo_4l);
                cbo_1.Show(); cbo_2.Show(); cbo_4l.Show();

                dtp_3frm.Value = DateTime.Parse(dtp_3frm.Value.ToString("yyyy-MM-01"));
                dtp_3frm.Show(); dtp_3to.Show();

                groupBox1.Height = 152;
                this.Height = 291;
            }

            this.Text = this.Text + " | " + rpt_code;
            isReady = true;
        }


        private void btn_proceed_Click(object sender, EventArgs e)
        {
            Boolean proceed = false;

            if (rpt_code == "M102" || rpt_code == "M103")
            {
                if (cbo_1.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select pay period.");
                    cbo_1.DroppedDown = true;
                    return;
                }
            }


            if (MessageBox.Show("Are you sure you want to proceed?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (rpt_code == "P101")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_payroll_summary((cbo_1.SelectedValue ?? "").ToString(), cbo_2frm.Text, (cbo_2frm.SelectedValue ?? "").ToString(), cbo_2to.Text, (cbo_2to.SelectedValue ?? "").ToString(), (cbo_3.SelectedValue ?? "").ToString(), cbo_4.Text, cbo_5.Text, chk_4.Checked, chk_5.Checked);

                    frm.ShowDialog();

                }
                else if (rpt_code == "P102")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_payslip((cbo_1.SelectedValue ?? "").ToString(), cbo_2frm.Text, (cbo_2frm.SelectedValue ?? "").ToString(), cbo_2to.Text, (cbo_2to.SelectedValue ?? "").ToString(), (cbo_3.SelectedValue ?? "").ToString(), cbo_4.Text, cbo_5.Text, chk_4.Checked, chk_5.Checked);

                    frm.ShowDialog();
                }
                else if (rpt_code == "P202")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_sss_contribution_summary(cbo_1.Text, cbo_2frm.Text, cbo_2to.Text, cbo_3.Text, (cbo_3.SelectedValue ?? "").ToString(), cbo_4l.Text, (cbo_4l.SelectedValue ?? "").ToString());

                    frm.ShowDialog();
                }
                else if (rpt_code == "P203")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_philhealth_contribution_summary(cbo_1.Text, cbo_2frm.Text, cbo_2to.Text, cbo_3.Text, (cbo_3.SelectedValue ?? "").ToString(), cbo_4l.Text, (cbo_4l.SelectedValue ?? "").ToString());

                    frm.ShowDialog();
                }

                else if (rpt_code == "E103")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_employee_listing(cbo_1.Text, (cbo_1.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_2.SelectedValue ?? "").ToString(), chk_3.Checked, chk_4.Checked);
                    frm.ShowDialog();
                }

                else if (rpt_code == "M102")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_other_earning_summary((cbo_1.SelectedValue ?? "").ToString(), cbo_1.Text, (cbo_2.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_3.SelectedValue ?? "").ToString(), cbo_3.Text);
                    frm.ShowDialog();
                }
                else if (rpt_code == "M103")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_other_deduction_summary((cbo_1.SelectedValue ?? "").ToString(), cbo_1.Text, (cbo_2.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_3.SelectedValue ?? "").ToString(), cbo_3.Text);
                    frm.ShowDialog();
                }
                else if (rpt_code == "M104")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_other_deduction_summary_per_employee((cbo_1.SelectedValue ?? "").ToString(), cbo_1.Text, (cbo_2.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_3.SelectedValue ?? "").ToString(), cbo_3.Text, (cbo_4l.SelectedValue ?? "").ToString());
                    frm.ShowDialog();
                }

                else if (rpt_code == "T101")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_employee_dtr((cbo_1.SelectedValue ?? "").ToString(), cbo_1.Text, (cbo_2.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_4l.SelectedValue ?? "").ToString(), cbo_4l.Text, dtp_3frm.Value.ToString("yyyy-MM-dd"), dtp_3to.Value.ToString("yyyy-MM-dd"));

                    frm.ShowDialog();
                }
                else if (rpt_code == "T102")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_daily_timelog((cbo_1.SelectedValue ?? "").ToString(), cbo_1.Text, (cbo_2.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_4l.SelectedValue ?? "").ToString(), cbo_4l.Text, dtp_3frm.Value.ToString("yyyy-MM-dd"), dtp_3to.Value.ToString("yyyy-MM-dd"));

                    frm.ShowDialog();
                }
                else if (rpt_code == "T103")
                {
                    RPT_RES_entry frm = new RPT_RES_entry(rpt_code, this.Text);
                    frm.print_absencelateundertime((cbo_1.SelectedValue ?? "").ToString(), cbo_1.Text, (cbo_2.SelectedValue ?? "").ToString(), cbo_2.Text, (cbo_4l.SelectedValue ?? "").ToString(), cbo_4l.Text, dtp_3frm.Value.ToString("yyyy-MM-dd"), dtp_3to.Value.ToString("yyyy-MM-dd"));

                    frm.ShowDialog();
                }
            }
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }



        private void cbo_typ_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isReady) return;

            if (rpt_code == "E103")
            {
                cbo_2.DroppedDown = true;
            }
        }

        private void RPT_RES_entry2_Load(object sender, EventArgs e)
        {

        }


        public void load_payroll_period(ComboBox cbo)
        {
            try
            {
                cbo.DataSource = db.QueryBySQLCode("SELECT pay_code, pay_code||' | '||to_char(date_from,'YYYY/MM/DD')||' - '||to_char(date_to,'YYYY/MM/DD') AS pay_period FROM rssys.hr_payrollpariod WHERE COALESCE(cancel,cancel,'')<>'Y' ORDER BY pay_code ASC");
                cbo.DisplayMember = "pay_period";
                cbo.ValueMember = "pay_code";
                cbo.SelectedIndex = -1;
            }
            catch { }
        }
        
        public void load_department(ComboBox cbo)
        {
            try
            {
                cbo.DataSource = db.QueryBySQLCode("SELECT deptid, dept_name FROM rssys.hr_department WHERE COALESCE(cancel,cancel,'')<>'Y' ORDER BY deptid ASC");
                cbo.DisplayMember = "dept_name";
                cbo.ValueMember = "deptid";
                cbo.SelectedIndex = -1;
            }
            catch { }
        }
        
        public void load_employee(ComboBox cbo)
        {
            try
            {
                cbo.DataSource = db.QueryBySQLCode("SELECT empid, firstname||' '||lastname||' '||mi AS empname  FROM rssys.hr_employee ORDER BY empid ASC");
                cbo.DisplayMember = "empname";
                cbo.ValueMember = "empid";
                cbo.SelectedIndex = -1;
            }
            catch { }
        }

        public void load_other_income(ComboBox cbo)
        {
            try
            {
                cbo.DataSource = db.QueryBySQLCode("SELECT code, description  FROM rssys.hr_other_earnings WHERE COALESCE(cancel,cancel,'')<>'Y' ORDER BY code ASC");
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch { }
        }
        public void load_other_deduction(ComboBox cbo)
        {
            try
            {
                cbo.DataSource = db.QueryBySQLCode("SELECT code, description  FROM rssys.hr_other_deductions WHERE COALESCE(cancel,cancel,'')<>'Y' ORDER BY code ASC");
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch { }
        }
    }
}
