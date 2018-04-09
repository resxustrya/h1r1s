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

    public partial class p_PayslipsEntry : Form
    {
        p_Payslips _p_frm_reg = null;
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
        Boolean isnew = false;
        String code = "";

        public p_PayslipsEntry()
        {
            InitializeComponent();
            gc.load_payroll_period(cbo_period_code);
            gc.load_dept(cbo_department_frm);
            gc.load_dept(cbo_department_until);
            gc.load_employee(cbo_employee);

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
        public p_PayslipsEntry(p_Payslips frm, String pk, Boolean _isnew)
        {
            InitializeComponent();
            gc.load_payroll_period(cbo_period_code);
            gc.load_dept(cbo_department_frm);
            gc.load_dept(cbo_department_until);
            gc.load_employee(cbo_employee);
            _p_frm_reg = frm;
            code = pk;
            isnew = _isnew;
            if (isnew == false)
            {
                init_load(code);
            }
        }
        public void init_load(String pk)
        {
            DataTable dt = null;
            dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_payslip WHERE payslip_code='" + pk + "'");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    txt_code.Text = dt.Rows[i]["payslip_code"].ToString();
                    cbo_period_code.SelectedValue = dt.Rows[i]["payroll_period_code"].ToString();
                    cbo_department_frm.SelectedValue = dt.Rows[i]["dept_frm"].ToString();
                    cbo_department_until.SelectedValue = dt.Rows[i]["dept_until"].ToString();
                    cbo_employee.SelectedValue = dt.Rows[i]["employee"].ToString();
                    cbo_report_type.Text = dt.Rows[i]["report_type"].ToString();
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void p_PayrollRegisterEntry_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            String payrollreg_code, dept_frm, dept_until, employee, report_type, payroll_period_code, payroll_period_desc;
            String col = "", val = "";
            Boolean success = false;
            String table = "hr_payslip";

            if (cbo_period_code.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select Period");
                cbo_period_code.DroppedDown = true;
            }
            else if (cbo_department_frm.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select FROM Department.");
                cbo_department_frm.DroppedDown = true;
            }
            else if (cbo_department_until.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select UNTIL Department.");
                cbo_department_until.DroppedDown = true;
            }
            else if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select Employee.");
                cbo_employee.DroppedDown = true;
            }
            else if (cbo_report_type.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select Report Type.");
                cbo_report_type.DroppedDown = true;
            }
            else
            {
                payrollreg_code = txt_code.Text;
                payroll_period_code = cbo_period_code.SelectedValue.ToString();
                payroll_period_desc = cbo_period_code.Text;
                dept_frm = cbo_department_frm.SelectedValue.ToString();
                dept_until = cbo_department_until.SelectedValue.ToString();
                employee = cbo_employee.SelectedValue.ToString();
                report_type = cbo_report_type.Text;


                if (isnew)
                {
                    payrollreg_code = db.get_pk("payslip_code");
                    col = "payslip_code,dept_frm,dept_until,employee,report_type,payroll_period_code,payroll_period_desc";
                    val = "'" + payrollreg_code + "', '" + dept_frm + "', '" + dept_until + "', '" + employee + "', '" + report_type + "', '" + payroll_period_code + "', '" + payroll_period_desc + "'";

                    if (db.InsertOnTable(table, col, val))
                    {
                        db.set_pkm99("payslip_code", db.get_nextincrementlimitchar(payrollreg_code, 8));
                        success = true;
                        //add_items(loan_code);
                    }
                    else
                    {
                        success = false;
                        db.DeleteOnTable(table, "payslip_code='" + payrollreg_code + "'");
                        MessageBox.Show("Failed on saving.");
                    }


                }
                else
                {
                    col = "payslip_code='" + payrollreg_code + "', dept_frm='" + dept_frm + "', dept_until='" + dept_until + "', employee='" + employee + "', report_type='" + report_type + "', payroll_period_code='" + payroll_period_code + "', payroll_period_desc='" + payroll_period_desc + "'";

                    if (db.UpdateOnTable(table, col, "payslip_code='" + payrollreg_code + "'"))
                    {
                        //db.DeleteOnTable("soalne", "loan_code='" + loan_code + "'");
                        //add_items(code);

                        success = true;
                    }
                    else
                    {
                        MessageBox.Show("Failed on saving.");
                        success = false;
                    }


                }
                if (success)
                {

                    this.Close();
                    frm_clear();

                }
            }
        }
        public void frm_clear()
        {
            cbo_period_code.SelectedIndex = -1;
            cbo_report_type.SelectedIndex = -1;
            cbo_employee.SelectedIndex = -1;
            cbo_department_frm.SelectedIndex = -1;
            cbo_department_until.SelectedIndex = -1;
            txt_code.Text = "";

        }
    }
}
