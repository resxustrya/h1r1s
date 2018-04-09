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
    public partial class p_GeneratePayroll : Form
    {

      
        private GlobalClass gc;
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
       
        public p_GeneratePayroll()
        {
            InitializeComponent();
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            bg_worker.RunWorkerAsync();
        }

        private void bg_worker_DoWork(object sender, DoWorkEventArgs e)
        {
            String table = "hr_emp_payroll", col = "", val = "", emp_pay_code = "";
            String summ_code = "", empid = "", days_worked = "", absences = "", late = "", undertime = "", overtime = "", ppid = "", sss_bracket= "", philhealth_bracket= "", pagibig_bracket = "";
            Double total_late = 0.00, total_under_time = 0.00, total_overtime = 0.00, leaves_amnt = 0.00, other_deduc = 0.00, other_earn = 0.00, advances_loans = 0.00, pag_ibig_a = 0.00, philhealth_cont_a = 0.00, sss_cont_a = 0.00;
            DataTable dtr = get_generated_dtr(), dttmp = new DataTable();
            Boolean success = false;
            int bar = 1;
            if (dtr.Rows.Count > 0)
            {
                try
                {
                    pbar.Invoke(new Action(() =>
                    {
                        pbar.Maximum = dtr.Rows.Count;
                    }));
                    for (int r = 0; r < dtr.Rows.Count; r++)
                    {
                        ppid = dtr.Rows[r]["ppid"].ToString();
                        empid = dtr.Rows[r]["empid"].ToString();
                        days_worked = dtr.Rows[r]["days_worked"].ToString();
                        absences = dtr.Rows[r]["absences"].ToString();
                        late = dtr.Rows[r]["late"].ToString();
                        undertime = dtr.Rows[r]["undertime"].ToString();
                        overtime = dtr.Rows[r]["total_overtime"].ToString();
                        try
                        {
                            total_late = Convert.ToDouble(TimeSpan.Parse(late).TotalHours);

                        }
                        catch (Exception ex)
                        {
                            total_late = 0.00;
                        }

                        try
                        {
                            total_under_time = Convert.ToDouble(TimeSpan.Parse(undertime).TotalHours);

                        }
                        catch (Exception ex)
                        {
                            total_under_time = 0.00;
                        }
                        try
                        {
                            total_overtime = Convert.ToDouble(TimeSpan.Parse(overtime).TotalHours);
                        }
                        catch (Exception ex)
                        {
                            total_overtime = 0.00;
                        }

                        chk_oearn.Invoke(new Action(() => {
                            if (chk_oearn.Checked)
                            {
                                DataTable tdt = db.QueryBySQLCode("SELECT * FROM rssys.hr_deduction_entry WHERE payroll_period='" + ppid + "' AND emp_no='" + empid + "'");
                                if (tdt != null)
                                {
                                    for (int i = 0; i < tdt.Rows.Count; i++)
                                    {
                                        other_earn += gm.toNormalDoubleFormat(tdt.Rows[i]["amount"].ToString());
                                    }
                                }
                            }
                        }));
                        chk_odeduc.Invoke(new Action(() =>
                        {
                            if (chk_odeduc.Checked)
                            {
                                DataTable tdt = db.QueryBySQLCode("SELECT * FROM rssys.hr_deduction_entry WHERE payroll_period='" + ppid + "' AND emp_no='" + empid + "'");

                                if (tdt != null)
                                {
                                    for (int i = 0; i < tdt.Rows.Count; i++)
                                    {
                                        other_deduc += gm.toNormalDoubleFormat(tdt.Rows[i]["amount"].ToString());
                                    }
                                }
                            }
                        }));


                        chk_loan.Invoke(new Action(() =>
                        {
                            if (chk_loan.Checked)
                            {
                                DataTable tdt = db.QueryBySQLCode("SELECT * FROM rssys.hr_loanhdr l LEFT JOIN rssys.hr_payrollpariod pp ON l.loan_transdate BETWEEN pp.date_from AND pp.date_to  WHERE pp.pay_code='" + ppid + "' AND l.employee_no='" + empid + "'");

                                if (tdt != null)
                                {
                                    for (int i = 0; i < tdt.Rows.Count; i++)
                                    {
                                        advances_loans += gm.toNormalDoubleFormat(tdt.Rows[i]["loan_amount"].ToString());
                                    }
                                }
                            }
                        }));

                        chk_leave.Invoke(new Action(() =>
                        {
                            if (chk_leave.Checked)
                            {
                                DataTable tdt = db.QueryBySQLCode("SELECT * from rssys.hr_leaves l LEFT JOIN rssys.hr_payrollpariod pp ON l.d_filed BETWEEN pp.date_from AND pp.date_to WHERE COALESCE(l.cancel,l.cancel,'')<>'Y' AND pp.pay_code='" + ppid + "' AND l.empid='" + empid + "'");

                                if (tdt != null)
                                {
                                    for (int i = 0; i < tdt.Rows.Count; i++)
                                    {
                                        leaves_amnt += gm.toNormalDoubleFormat(tdt.Rows[i]["leave_amount"].ToString());
                                    }
                                }
                            }
                        }));
                        

                        emp_pay_code = db.get_pk("emp_pay_code");
                        col = "emp_pay_code,empid,days_worked,abcences,late,undertime,overtime,ppid,other_deduction, other_earnings, leave_amnt, advances_loans";
                        val = "'" + emp_pay_code + "','" + empid + "','" + days_worked + "','" + absences + "','" + total_late.ToString("0.00") + "','" + total_under_time.ToString("0.00") + "','" + total_overtime.ToString("0.00") + "','" + ppid + "','" + other_deduc + "','" + other_earn + "','" + leaves_amnt + "','" + advances_loans + "'";



                        dttmp = db.QueryBySQLCode("SELECT  COALESCE(emp_ee,0.00) +  COALESCE(emp_er,0.00) AS total, pagibig_bracket FROM rssys.hr_hdmf h JOIN rssys.hr_employee e ON (e.pagibig_bracket=h.code AND e.empid='" + empid + "')");
                        if (dttmp != null)
                        {
                            try
                            {
                                pag_ibig_a = gm.toNormalDoubleFormat(dttmp.Rows[0]["total"].ToString());
                                pagibig_bracket = dttmp.Rows[0]["pagibig_bracket"].ToString();
                            } catch { }
                        }
                        dttmp = db.QueryBySQLCode("SELECT COALESCE(empshare_sc,0.00) + COALESCE(empshare_ec,0.00) + COALESCE(s_ec,0.00) AS total, sss_bracket FROM rssys.hr_sss h JOIN rssys.hr_employee e ON (e.sss_bracket=h.code AND e.empid='" + empid + "')");
                        if (dttmp != null)
                        {
                            try
                            {
                                sss_cont_a = gm.toNormalDoubleFormat(dttmp.Rows[0]["total"].ToString());
                                sss_bracket = dttmp.Rows[0]["sss_bracket"].ToString();
                            } catch { }
                        }
                        dttmp = db.QueryBySQLCode("SELECT COALESCE(emp_ee,0.00) + COALESCE(emp_er,0.00) AS total, philhealth_bracket FROM rssys.hr_philhealth h JOIN rssys.hr_employee e ON (e.philhealth_bracket=h.code AND e.empid='" + empid + "')");
                        if (dttmp != null)
                        {
                            try
                            {
                                philhealth_cont_a = gm.toNormalDoubleFormat(dttmp.Rows[0]["total"].ToString());
                                philhealth_bracket = dttmp.Rows[0]["philhealth_bracket"].ToString();
                            } catch { }
                        }

                        col += ",regular_pay, basic_pay, vl_a, vl_b, sl_a, sl_b, pl_a, pl_b, regular_ot_a, reqular_ot_b, dayoff_ot_a, dayoff_ot_b, legal_hol_ot_a, legal_hol_ot_b, special_hol_ot_a, special_hol_ot_b, legal_hol_pay_a, legal_hol_pay_b, spl_hol_pay_a, spl_hol_pay_b,night_diff_a, night_diff_b, sss_cont_a, sss_cont_b,  sss_cont_c, philhealth_cont_a, philhealth_cont_b, pag_ibig_a, pag_ibig_b, others, w_tax, late_amnt, absent_amnt, sss_bracket, philhealth_bracket, pagibig_bracket";
                        val += ",'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0','0', '" + sss_cont_a.ToString("0.00") + "', '0',  '0', '" + philhealth_cont_a.ToString("0.00") + "', '0', '" + pag_ibig_a.ToString("0.00") + "', '0', '0', '0', '0', '0','" + sss_bracket + "','" + philhealth_bracket + "','" + pagibig_bracket + "'";
                        
                        if (db.InsertOnTable(table, col, val))
                        {
                            col = "isgenerated='1'";

                            if (db.UpdateOnTable("hr_dtr_sum_employees", col, "empid='" + empid + "' AND ppid='" + ppid + "'"))
                            {
                                success = true;
                                inc_pbar(bar, dtr.Rows.Count);
                                bar++;
                            }
                            db.set_pkm99("emp_pay_code", db.get_nextincrementlimitchar(emp_pay_code, 8));
                        }
                    }
                    if (success)
                    {
                        String period = get_payrol_period(ppid);
                        MessageBox.Show("New payroll was generated From " + period + " .");
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Payroll cannot be generated. Something went wrong. " + ex.Message);
                }

            }
            else
            {
                MessageBox.Show("No generated DTR is available.");

            }
        }

        private void inc_pbar(int i, int rw)
        {
            try
            {

                if (pbar.Value <= rw)
                {
                    pbar.Invoke(new Action(() =>
                    {
                        pbar.Value = i;
                    }));

                }
                else
                {
                    pbar.Invoke(new Action(() =>
                    {
                        pbar.Value = rw;
                    }));
                }

            }
            catch (Exception)
            {

            }
        }
        public void disp_list()
        {
            DataTable dt = db.QueryBySQLCode("");

        }
        private DataTable get_generated_dtr()
        {
            DataTable dt = null;
            try
            {
                dt = db.QueryBySQLCode("SELECT * from rssys.hr_dtr_sum_employees WHERE isgenerated = '0'");
            }
            catch { }

            return dt;
        }

        private String get_payrol_period(String ppid)
        {
            String period = "";
            DataTable dt = null;
            try
            {
                dt = db.QueryBySQLCode("SELECT concat(to_char(date_from, 'mm/dd/yyyy'),' To ',to_char(date_to, 'mm/dd/yyyy')) as period FROM rssys.hr_payrollpariod WHERE pay_code='" + ppid + "'");
            }
            catch { }
            if(dt.Rows.Count > 0)
            {
                period = dt.Rows[0]["period"].ToString();
            }
            return period;
        }

        private void p_GeneratePayroll_Load(object sender, EventArgs e)
        {
            disp_list_history();
        }


        private void disp_list_history()
        {
            dgv_list.Rows.Clear();
            try
            {

                DataTable dt = db.QueryBySQLCode("SELECT sh.*, concat(firstname,' ',lastname) AS employee FROM rssys.hr_dtr_sum_hdr sh LEFT JOIN rssys.hr_employee e ON e.empid=sh.empid  ORDER BY date_generated DESC, time_generated DESC");

                if (dt.Rows.Count > 0)
                {
                    for (int r = 0; dt.Rows.Count > r; r++)
                    {
                        int i = dgv_list.Rows.Add();
                        DataGridViewRow row = dgv_list.Rows[i];

                        row.Cells["dgvl_date"].Value = gm.toDateString(dt.Rows[r]["date_generated"].ToString(), "");
                        row.Cells["dgvl_time"].Value = dt.Rows[r]["time_generated"].ToString();
                        row.Cells["dgvl_payroll"].Value = gm.toDateString(dt.Rows[r]["date_from"].ToString(), "") + " TO " + gm.toDateString(dt.Rows[r]["date_to"].ToString(), "");

                        row.Cells["dgvl_userid"].Value = dt.Rows[r]["empid"].ToString();
                        //cbo_employee.SelectedValue = dt.Rows[r]["empid"].ToString();
                        row.Cells["dgvl_employee"].Value = dt.Rows[r]["employee"].ToString();

                    }
                }
            }
            catch { }

        }
    }
}
