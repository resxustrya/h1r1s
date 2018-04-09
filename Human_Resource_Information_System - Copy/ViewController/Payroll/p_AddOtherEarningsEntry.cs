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
    public partial class p_AddOtherEarningsEntry : Form
    {
        GlobalClass gc = new GlobalClass();
        GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
        Boolean isnew = false;
        String code = "";
        p_OtherEarningEntry _frm_earn = null;
        public p_AddOtherEarningsEntry()
        {
            InitializeComponent();
            gc.load_employee(cbo_employee);
        }
        public p_AddOtherEarningsEntry(p_OtherEarningEntry frm, String pk, Boolean _isnew)
        {
            InitializeComponent();
           
            gc.load_employee(cbo_employee);
            _frm_earn = frm;
            code = pk;
            isnew = _isnew;
            if(isnew == false)
            {
                init_load(code);
            }
        }
        public void init_load(String pk)
        {
            DataTable dt = null;
            dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_earning_entry WHERE   entcode='" + pk + "'");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    txt_code.Text = dt.Rows[i]["entcode"].ToString();
                    cbo_employee.SelectedValue = dt.Rows[i]["emp_no"].ToString();
                    txt_amount.Text = dt.Rows[i]["amount"].ToString();
                   
                }
            }
        }
        private void p_AddOtherEarningsEntry_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            String entcode, payroll_period, emp_no, emp_name, earning_code, amount;
            String col = "", val = "";
            Boolean success = false;
            String table = "hr_earning_entry";

            if (cbo_employee.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select Employee");
                cbo_employee.DroppedDown = true;
            }
            else if (_frm_earn.cbo_payroll_period.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select Payroll Period.");
                _frm_earn.cbo_payroll_period.DroppedDown = true;

            }
            else if (_frm_earn.cbo_earnings_code.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select Earnings.");
                _frm_earn.cbo_earnings_code.DroppedDown = true;
            }
            else if (txt_amount.Text == "")
            {
                MessageBox.Show("Please Enter Amount.");
                txt_amount.Focus();
            }
            else
            {
                entcode = txt_code.Text;
                payroll_period = _frm_earn.cbo_payroll_period.SelectedValue.ToString();
                emp_no = cbo_employee.SelectedValue.ToString();
                emp_name = cbo_employee.Text;
                earning_code = _frm_earn.cbo_earnings_code.SelectedValue.ToString();
                amount = txt_amount.Text;
               


                if (isnew)
                {
                    entcode = db.get_pk("entcode");
                    col = "entcode,payroll_period,emp_no,emp_name,earning_code,amount";
                    val = "'" + entcode + "', '" + payroll_period + "', '" + emp_no + "', '" + emp_name + "', '" + earning_code + "', '" + amount  + "'";

                    if (db.InsertOnTable(table, col, val))
                    {
                        db.set_pkm99("entcode", db.get_nextincrementlimitchar(entcode, 8));
                        success = true;
                        //add_items(loan_code);
                    }
                    else
                    {
                        success = false;
                        db.DeleteOnTable(table, "entcode='" + entcode + "'");
                        MessageBox.Show("Failed on saving.");
                    }


                }
                else
                {
                    
                    col = "entcode='" + entcode + "', payroll_period='" + payroll_period + "', emp_no='" + emp_no + "', emp_name='" + emp_name + "', earning_code='" + earning_code + "', amount='" + amount + "'";

                    if (db.UpdateOnTable(table, col, "entcode='" + entcode + "'"))
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
        }//end of method

        public void frm_clear()
        {
            txt_code.Text = "";
            cbo_employee.SelectedIndex = -1;
            txt_amount.Text = "";
            
        }
    }
}
