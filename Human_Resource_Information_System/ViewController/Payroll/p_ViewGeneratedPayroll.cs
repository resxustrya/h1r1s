using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace Human_Resource_Information_System
{
    public partial class p_ViewGeneratedPayroll : Form
    {
        Boolean seltbp = false;
        private Boolean isnew = false;
        private GlobalClass gc = new GlobalClass();
        private GlobalMethod gm = new GlobalMethod();
        thisDatabase db = new thisDatabase();
        private Double daily_rate = 0.00, pay_rate = 0.00, hour_rate = 0.00, minute_rate = 0.00, gross_pay = 0.00, basic_pay = 0.00, monthly_rate = 0.00, late_amt = 0.00, absent_amt = 0.00, lat_ut = 0.00, reg_ot_a = 0.00, reg_ot_b_total = 0.00;

        int count_dayoff = 0, dayoff = 0, days = 0;
        public p_ViewGeneratedPayroll()
        {
            InitializeComponent();
        }

        private void p_ViewGeneratedPayroll_Load(object sender, EventArgs e)
        {
            disp_payroll();
        }

        private void goto_win2()
        {
            seltbp = true;
            tbcntrl_option.SelectedTab = tpg_option_2;
            tpg_info.Show();

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

        private void btn_upd_Click(object sender, EventArgs e)
        {

            int r = -1;
            String code = "", name = "",empid = "",ppid="";
            try
            {
                if (dgv_list.Rows.Count > 1)
                {
                    r = dgv_list.CurrentRow.Index;

                    try
                    {
                        code = dgv_list["pay_code", r].Value.ToString();
                        empid = dgv_list["empid", r].Value.ToString();
                        ppid = dgv_list["ppid", r].Value.ToString();
                        lbl_code.Text = code;
                        display_emp_payroll(code, empid, ppid);
                        goto_win2();
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Select employee payroll.");
                    }

                    
                }
                else
                {
                    MessageBox.Show("Payroll list is empty.");
                }
            }
            catch (Exception ex)
            {

            }

        }
        private void disp_payroll()
        {
            try { dgv_list.Rows.Clear(); }
            catch (Exception) { }
            int i = 0;
            String query = "SELECT p.pay_code,ep.emp_pay_code,emp.empid,CONCAT(emp.firstname, ' ',emp.lastname) as name, CONCAT(to_char(p.date_from, 'mm/dd/yyyy'),' To ',to_char(p.date_to, 'mm/dd/yyyy')) as period FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee emp ON emp.empid = ep.empid LEFT JOIN rssys.hr_payrollpariod p ON p.pay_code = ep.ppid";
            if (txt_search.Text != "")
            {
                query += " WHERE emp.firstname LIKE '"+txt_search.Text+"%' OR emp.lastname LIKE '" +txt_search.Text+ "%'";
            }
            query += " ORDER BY p.date_from,p.date_to,ep.emp_pay_code ASC";
            
            try
            {
                DataTable dt = db.QueryBySQLCode(query);

                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    i = dgv_list.Rows.Add();
                    DataGridViewRow row = dgv_list.Rows[i];

                    row.Cells["pay_code"].Value = dt.Rows[r]["emp_pay_code"].ToString();
                    row.Cells["empname"].Value = dt.Rows[r]["name"].ToString();
                    row.Cells["pay_period"].Value = dt.Rows[r]["period"].ToString();
                    row.Cells["empid"].Value = dt.Rows[r]["empid"].ToString();
                    row.Cells["ppid"].Value = dt.Rows[r]["pay_code"].ToString();
                    i++;
                }
            }
            catch { }
            
        }

     
        private void dgv_list_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void display_emp_payroll(String pay_code, String empid,String ppid)
        {
          
            DataTable payroll = null;
               
            try
            {
                payroll = db.QueryBySQLCode("SELECT ep.*,emp.fixed_rate,emp.empid,emp.dayoff1,emp.dayoff2,emp.shift_sched_from,emp.shift_sched_to,emp.pay_rate,rt.*,CONCAT(emp.firstname, ' ',emp.lastname) as name, CONCAT(to_char(p.date_from, 'mm/dd/yyyy'),' To ',to_char(p.date_to, 'mm/dd/yyyy')) as period FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee emp ON emp.empid = ep.empid LEFT JOIN rssys.hr_payrollpariod p ON p.pay_code = ep.ppid LEFT JOIN rssys.hr_rate_type rt ON rt.ratecode = emp.rate_type WHERE emp.empid = '" + empid + "' and ep.emp_pay_code = '" + pay_code + "' ORDER BY p.date_from,p.date_to ASC  ");
                if(payroll.Rows.Count > 0)
                {

                    try
                    {

                    
                        disp_legal_hol_pay(empid, ppid);
                        txt_empname.Text = payroll.Rows[0]["name"].ToString();
                        txt_pay_period.Text = (payroll.Rows[0]["period"]??"0").ToString();

                        //Convert.ToDouble(dt.Rows[0]["pay_rate"]).ToString("N", new CultureInfo("en-US")); }

                        txt_pay_rate.Text = (payroll.Rows[0]["pay_rate"]??"0.00").ToString();
                        txt_rate_type.Text = (payroll.Rows[0]["description"]??"0.00").ToString();
                        txt_dayswoked.Text = Convert.ToDouble((payroll.Rows[0]["days_worked"]??"0.00").ToString()).ToString("0.00");
                        txt_absent.Text = Convert.ToDouble((payroll.Rows[0]["abcences"]??"0").ToString()).ToString("0.00");
                        
                        if (payroll.Rows[0]["late"].ToString() == "" && payroll.Rows[0]["undertime"].ToString() == "")
                        {
                            txt_late_ut.Text = "0.00";
                        }
                        else if (payroll.Rows[0]["late"].ToString() == "" && payroll.Rows[0]["undertime"].ToString() != "")
                        {
                            txt_late_ut.Text = payroll.Rows[0]["undertime"].ToString();
                        }
                        else if (payroll.Rows[0]["undertime"].ToString() == ""  && payroll.Rows[0]["late"].ToString() != "")
                        {
                            txt_late_ut.Text = payroll.Rows[0]["late"].ToString();
                        } else
                        {
                            txt_late_ut.Text = (Convert.ToDouble(payroll.Rows[0]["late"].ToString()) + Convert.ToDouble(payroll.Rows[0]["undertime"].ToString())).ToString("0.00");
                        }
                        
                    }catch(Exception ex)
                    {

                    }

                    pay_rate = Convert.ToDouble((payroll.Rows[0]["pay_rate"] ?? "0.00").ToString());

                    if (payroll.Rows[0]["dayoff1"].ToString() != "-1")
                    {
                        count_dayoff = 1;
                    }

                    if (payroll.Rows[0]["dayoff2"].ToString() != "-1")
                    {
                        count_dayoff = 2;
                    }

                    if (count_dayoff == 2)
                    {
                        days = 261;
                    }
                    else if (count_dayoff == 1)
                    {
                        days = 312;
                    }
                    else
                    {
                        days = 365;
                    }

                    daily_rate = (pay_rate * 12) / days; //pay_rate = monthly rate
                    hour_rate = (daily_rate / 8);
                    minute_rate = (hour_rate / 60);

                    txt_min_rate.Text = minute_rate.ToString("N", new CultureInfo("en-US"));
                    txt_hourly_rate.Text = hour_rate.ToString("N", new CultureInfo("en-US"));
                    txt_daily_rate.Text = daily_rate.ToString("N", new CultureInfo("en-US"));
                    

                    lat_ut = (Convert.ToDouble(txt_late_ut.Text) * 60) * minute_rate;
                    txt_late_ut_amt.Text = lat_ut.ToString("0.00");
                    txt_reg_ot_a.Text = (payroll.Rows[0]["regular_ot_a"] ?? "0.00").ToString();
                    // reg_ot_a = Convert.ToDouble(txt_reg_ot_a.Text);

                    if (payroll.Rows[0]["ratecode"].ToString() == "M")
                    {
                        if (payroll.Rows[0]["fixed_rate"].ToString() == "1")
                        {
                            txt_regpay.Text = pay_rate.ToString("0.00");
                            absent_amt = (Convert.ToDouble(txt_absent.Text) * daily_rate);
                            txt_absent_amount.Text = absent_amt.ToString("N", new CultureInfo("en-US"));
                            //pay_rate = pay_rate / 2; //pay_rate = half of the month rate
                            late_amt = Convert.ToDouble(txt_late_ut_amt.Text);
                            basic_pay = (pay_rate - (late_amt + absent_amt));
                            txt_reg_ot_b.Text = reg_ot_b_total.ToString("0.00");
                            txt_basic_pay.Text = basic_pay.ToString("0.00");

                        }
                        else
                        {
                            try
                            {
                                txt_regpay.Text = (Convert.ToDouble(txt_dayswoked.Text) * daily_rate).ToString("0.00");
                            }
                            catch (Exception ex)
                            {
                                txt_regpay.Text = "0.00";
                            }
                        }
                    }
                    else if (payroll.Rows[0]["ratecode"].ToString() == "D")
                    {
                        if (payroll.Rows[0]["fixed_rate"].ToString() == "1")
                        {
                            
                            txt_regpay.Text = pay_rate.ToString("0.00");
                            absent_amt = (Convert.ToDouble(txt_absent.Text) * daily_rate);
                            txt_absent_amount.Text = absent_amt.ToString("N", new CultureInfo("en-US"));
                            //pay_rate = pay_rate / 2; //pay_rate = half of the month rate
                            late_amt = Convert.ToDouble(txt_late_ut_amt.Text);
                            basic_pay = (Convert.ToDouble(txt_dayswoked.Text) * daily_rate) - late_amt;
                            txt_reg_ot_b.Text = reg_ot_b_total.ToString("0.00");
                            txt_basic_pay.Text = basic_pay.ToString("0.00");
                        }
                    }
                    else if (payroll.Rows[0]["ratecode"].ToString() == "W")
                    {
                        try
                        {
                            int year = DateTime.Now.Year;
                            int week_days = 0;
                            if (DateTime.IsLeapYear(year))
                            {
                                if (count_dayoff == 2)
                                {
                                    week_days = 261;
                                }
                            }
                            else
                            {
                                week_days = 260;
                            }
                            daily_rate = (pay_rate * 12) / week_days;
                            txt_regpay.Text = (Convert.ToDouble(txt_dayswoked.Text) * daily_rate).ToString("0.00");
                        }
                        catch (Exception ex)
                        {
                            txt_regpay.Text = "0.00";
                        }
                    }



                    txt_vl_a.Text = (payroll.Rows[0]["vl_a"]??"0.00").ToString();
                    txt_vl_b.Text =  (payroll.Rows[0]["vl_b"]??"0.00").ToString();
                    txt_sl_a.Text = (payroll.Rows[0]["sl_a"]??"0.00").ToString();
                    txt_sl_b.Text = (payroll.Rows[0]["sl_b"]??"0.00").ToString();
                    txt_pl_a.Text = (payroll.Rows[0]["pl_a"]??"0.00").ToString();
                    txt_pl_b.Text = (payroll.Rows[0]["pl_b"]??"0.00").ToString();
                    
                    txt_dayoffot_a.Text = (payroll.Rows[0]["dayoff_ot_a"]??"0.00").ToString();
                    //txt_dayoffot_b.Text = (payroll.Rows[0]["dayoff_ot_b"]??"0.00").ToString();
                    txt_legalhol_ot_a.Text = (payroll.Rows[0]["legal_hol_ot_a"]??"0.00").ToString();
                   // txt_legalhol_ot_b.Text = (payroll.Rows[0]["legal_hol_ot_b"]??"0.00").ToString();
                    txt_specialhol_ot_a.Text = (payroll.Rows[0]["special_hol_ot_a"]??"0.00").ToString();
                   // txt_specialhol_ot_b.Text = (payroll.Rows[0]["special_hol_ot_b"]??"0.00").ToString();
                    txt_legalhol_pay_a.Text = (payroll.Rows[0]["legal_hol_pay_a"]??"0.00").ToString();
                    //txt_legalhol_pay_b.Text = (payroll.Rows[0]["legal_hol_pay_b"]??"0.00").ToString();
                    txt_specialhol_pay_a.Text = (payroll.Rows[0]["spl_hol_pay_a"]??"0.00").ToString();
                    txt_specialhol_pay_b.Text = (payroll.Rows[0]["spl_hol_pay_b"]??"0.00").ToString();
                    txt_night_diff_a.Text = (payroll.Rows[0]["night_diff_a"]??"0.00").ToString();
                    txt_night_diff_b.Text = (payroll.Rows[0]["night_diff_b"] ?? "0.00").ToString();
                    txt_other_earning.Text = (payroll.Rows[0]["other_earnings"]??"0.00").ToString();
                    txt_sss_a.Text = (payroll.Rows[0]["sss_cont_a"]??"0.00").ToString();
                    //txt_sss_b.Text = (payroll.Rows[0]["sss_cont_b"]??"0.00").ToString();
                    //txt_sss_con_c.Text = (payroll.Rows[0]["sss_cont_c"]??"0.00").ToString();
                    //txt_philhealth_a.Text = (payroll.Rows[0]["philhealth_cont_a"]??"0.00").ToString();
                    //txt_philhealth_b.Text = (payroll.Rows[0]["philhealth_cont_b"]??"0.00").ToString();
                    //txt_pagibig_a.Text = (payroll.Rows[0]["pag_ibig_a"]??"0.00").ToString();
                    //txt_pagibig_b.Text = payroll.Rows[0]["pag_ibig_b"].ToString();
                    txt_wtax.Text = (payroll.Rows[0]["w_tax"]??"0.00").ToString();
                    txt_other_deductions.Text = (payroll.Rows[0]["other_deduction"]??"0.00").ToString();
                    txt_others.Text = (payroll.Rows[0]["others"]??"0.00").ToString();
                    txt_advance_loans.Text = (payroll.Rows[0]["advances_loans"]??"0.00").ToString();

                    

                    
        
                }
                calculate_gross();
                calculate_net();
            }
            catch { }
        }

       
      
        private void btn_mainexit_Click(object sender, EventArgs e)
        {
            goto_win1();
            clear_field();
        }

        private void clear_field()
        {
            txt_absent.Text = "0.00";
            txt_dayswoked.Text = "0.00";
            txt_empname.Text = "--";
            txt_late_ut.Text = "0.00";
            txt_pay_period.Text = "0.00";
            txt_pay_rate.Text = "0.00";
            txt_rate_type.Text = "0.00";
            lbl_total_gross.Text = "0.00";
            lbl_total_tax.Text = "0.00";
            lbl_total_net.Text="0.00";
        }
        private void disp_legal_hol_pay(String empid, String ppid)
        {
            int sp = 0, lg = 0;
            String date = "";
            String time = "";
            String late = "";
            String ut = "";
            String timein = "";
            String timeout = "",time_log_in = "",time_log_out = "";
            TimeSpan total_late = new TimeSpan(0, 0, 0, 0, 0);
            TimeSpan total_ut = new TimeSpan(0, 0, 0, 0, 0);
            DataTable logs = null;
            DateTime d_late = new DateTime();
            try
            {
                DataTable period = db.QueryBySQLCode("SELECT date_from,date_to FROM rssys.hr_payrollpariod WHERE pay_code ='" + ppid + "'");
                String date_from = DateTime.Parse(period.Rows[0]["date_from"].ToString()).ToShortDateString();
                String date_to = DateTime.Parse(period.Rows[0]["date_to"].ToString()).ToShortDateString();       

                DataTable holiday = db.QueryBySQLCode("SELECT date_holiday FROM rssys.hr_holidays WHERE date_holiday BETWEEN '" + date_from + "' AND '" + date_to + "'");

                

                for(int r = 0; r < holiday.Rows.Count; r++)
                {
                    try
                    {
                        date = DateTime.Parse(holiday.Rows[r]["date_holiday"].ToString()).ToShortDateString();

                        logs = db.QueryBySQLCode("SELECT DISTINCT e.empid,e.shift_sched_from,e.shift_sched_to,work_date, (SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid WHERE t.work_date = '" + date + "' AND t.empid = '" + empid + "' ORDER BY work_date");

                        time_log_in = logs.Rows[0]["timein"].ToString();

                        if (!String.IsNullOrEmpty(time_log_in))
                        {
                          
                            timein = logs.Rows[0]["shift_sched_from"].ToString();

                            DateTime datetime_in = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + time_log_in);

                            DateTime datetime_from = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timein);

                            int res = DateTime.Compare(datetime_from, datetime_in);

                            if (res < 0)
                            {
                                TimeSpan diff = datetime_in.Subtract(datetime_from);
                                //MessageBox.Show("Time Log : " + time_log_in + " Time In : " + timein + "Late : " + diff.ToString());
                                total_late = total_ut + diff;
                                //MessageBox.Show("Total Late : " + total_ut);
                            }

                        }


                        time_log_out = logs.Rows[0]["timeout"].ToString();
                        if (!String.IsNullOrEmpty(time_log_out))
                        {
                            timeout = logs.Rows[0]["shift_sched_to"].ToString();
                            DateTime datetime_out = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + time_log_out);
                            DateTime datetime_to = Convert.ToDateTime(DateTime.Now.ToString("M/d/yyyy") + " " + timeout);
                            int res1 = DateTime.Compare(datetime_to, datetime_out);

                            if (res1 > 0)
                            {
                                TimeSpan diff = datetime_to.Subtract(datetime_out);
                                //MessageBox.Show("Out Time : " + datetime_to + " Time Out : " + datetime_out + " Undertime : " + diff);
                                total_ut = total_ut + diff;
                                //MessageBox.Show("Total Undertime : " + total_ut);
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

        }

        private void btn_mainsave_Click(object sender, EventArgs e)
        {

            
            String table = "hr_emp_payroll";
            Boolean success = false;
            String code = "";
            String col = "", days_worked = "", regular_pay = "", abcences = "", late = "", undertime = "", basic_pay = "", vl_a = "", vl_b = "", sl_a = "", sl_b = "", pl_a = "", pl_b = "", regular_ot_a = "", reqular_ot_b = "", dayoff_ot_a = "", dayoff_ot_b = "", legal_hol_ot_a = "", legal_hol_ot_b = "", special_hol_ot_a = "", special_hol_ot_b = "", legal_hol_pay_a = "", legal_hol_pay_b = "", spl_hol_pay_a = "", spl_hol_pay_b = "", night_diff_a = "", night_diff_b = "", other_earnings = "", sss_cont_b = "", sss_cont_a = "", sss_cont_c = "", philhealth_cont_a = "", philhealth_cont_b = "", pag_ibig_a = "", pag_ibig_b = "", other_deduction = "", others = "", advances_loans = "", empid = "", overtime = "", ppid = "", w_tax="",netpay = "";
            String val = "";

            code = lbl_code.Text != "" ? lbl_code.Text : "0.00";
            days_worked = txt_dayswoked.Text != "" ? txt_dayswoked.Text : "0.00";
            regular_pay = txt_regpay.Text != "" ? txt_regpay.Text : "0.00";
            abcences = txt_absent.Text != "" ? txt_absent.Text : "0.00";
            late = txt_late_ut.Text != "" ? txt_late_ut.Text : "0.00";
            basic_pay = txt_basic_pay.Text != "" ? txt_basic_pay.Text : "0.00";
            vl_a = txt_vl_a.Text != "" ? txt_vl_a.Text : "0.00";
            vl_b = txt_vl_b.Text != "" ? txt_vl_b.Text : "0.00";
            sl_a = txt_sl_a.Text != "" ? txt_sl_a.Text : "0.00";
            sl_b = txt_sl_b.Text != "" ? txt_sl_b.Text : "0.00";
            pl_a = txt_pl_a.Text != "" ? txt_pl_a.Text : "0.00";
            pl_b = txt_pl_b.Text != "" ? txt_pl_b.Text : "0.00";
            regular_ot_a = txt_reg_ot_a.Text != "" ? txt_reg_ot_a.Text : "0.00";
            reqular_ot_b = txt_reg_ot_b.Text != "" ? txt_reg_ot_b.Text : "0.00";
            dayoff_ot_a = txt_dayoffot_a.Text != "" ? txt_dayoffot_a.Text : "0.00";
            dayoff_ot_b = txt_dayoffot_b.Text != "" ? txt_dayoffot_b.Text : "0.00";
            legal_hol_ot_a = txt_legalhol_ot_a.Text != "" ? txt_legalhol_ot_a.Text : "0.00";
            legal_hol_ot_b = txt_legalhol_ot_b.Text != "" ? txt_legalhol_ot_b.Text : "0.00";
            special_hol_ot_a = txt_specialhol_ot_a.Text != "" ? txt_specialhol_ot_a.Text : "0:00";
            special_hol_ot_b = txt_specialhol_ot_b.Text != "" ? txt_specialhol_ot_b.Text : "0:00";
            legal_hol_pay_a = txt_legalhol_pay_a.Text != "" ? txt_legalhol_pay_a.Text : "0.00";
            legal_hol_pay_b = txt_legalhol_pay_b.Text != "" ? txt_legalhol_pay_b.Text : "0.00";
            spl_hol_pay_a = txt_specialhol_pay_a.Text != "" ? txt_specialhol_pay_a.Text : "0.00";
            spl_hol_pay_b = txt_specialhol_pay_b.Text != "" ? txt_specialhol_pay_b.Text : "0.00";
            night_diff_a = txt_night_diff_a.Text != "" ? txt_night_diff_a.Text : "0.00";
            night_diff_b = txt_night_diff_b.Text != "" ? txt_night_diff_b.Text : "0.00";
            other_earnings = txt_other_earning.Text != "" ? txt_other_earning.Text : "0.00";
            sss_cont_a = txt_sss_a.Text != "" ? txt_sss_a.Text : "0.00";
            sss_cont_b = "0.00";
            sss_cont_c = "0.00";
            philhealth_cont_a = txt_philhealth_a.Text != "" ? txt_philhealth_a.Text : "0.00";
            philhealth_cont_b = "0.00";
            pag_ibig_a = txt_pagibig_a.Text != "" ? txt_pagibig_a.Text : "0.00";
            pag_ibig_b = "0.00";
            w_tax = txt_wtax.Text != "" ? txt_wtax.Text : "0.00";
            other_deduction = txt_other_deductions.Text != "" ? txt_other_deductions.Text : "0.00";
            others = txt_others.Text != "" ? txt_others.Text : "0.00";
            advances_loans = txt_advance_loans.Text != "" ? txt_advance_loans.Text : "0.00";
            netpay = lbl_total_net.Text != "" ? lbl_total_net.Text.Replace(",","") : "0.00";


            col = "days_worked='" + days_worked + "',regular_pay='" + regular_pay + "',abcences='" + abcences + "',late='" + late + "',undertime='" + undertime + "',basic_pay='" + basic_pay + "',vl_a='" + vl_a + "',vl_b='" + vl_b + "',sl_a='" + sl_a + "',sl_b='" + sl_b + "',pl_a='" + pl_a + "',pl_b='" + pl_b + "',regular_ot_a='" + regular_ot_a + "',reqular_ot_b='" + reqular_ot_b + "',dayoff_ot_a='" + dayoff_ot_a + "',dayoff_ot_b='" + dayoff_ot_b + "',legal_hol_ot_a='" + legal_hol_ot_a + "',legal_hol_ot_b='" + legal_hol_ot_b + "',special_hol_ot_a='" + special_hol_ot_a + "',special_hol_ot_b='" + special_hol_ot_b + "',legal_hol_pay_a='" + legal_hol_pay_a + "',legal_hol_pay_b='" + legal_hol_pay_b + "',spl_hol_pay_a='" + spl_hol_pay_a + "',spl_hol_pay_b='" + spl_hol_pay_b + "',night_diff_a='" + night_diff_a + "',night_diff_b='" + night_diff_b + "',other_earnings='" + other_earnings + "',sss_cont_b='" + sss_cont_b + "',sss_cont_a='" + sss_cont_a + "',sss_cont_c='" + sss_cont_c + "',philhealth_cont_a='" + philhealth_cont_a + "',philhealth_cont_b='" + philhealth_cont_b + "',pag_ibig_a='" + pag_ibig_a + "',pag_ibig_b='" + pag_ibig_b + "',other_deduction='" + other_deduction + "',others='" + others + "',advances_loans='" + advances_loans + "',overtime='" + overtime + "',w_tax='" + w_tax + "',netpay = '" + netpay + "'";
           
            if (db.UpdateOnTable(table, col, "emp_pay_code='" + code + "'"))
            {
                success = true;
            }
            else
            {
                success = false;
                MessageBox.Show("Failed on saving.");
            }
            if(success)
            {
                goto_win1();

            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void clear_form()
        {
            Action<Control.ControlCollection> func = null;

            func = (controls) =>
            {
                foreach (Control control in controls)
                {
                    if (control is TextBox)
                    {
                        (control as TextBox).Text = "0.00";
                    }
                    else
                    {
                        func(control.Controls);
                    }
                    if (control is ComboBox)
                    {
                        (control as ComboBox).SelectedIndex = -1;
                    }
                    else
                    {
                        func(control.Controls);
                    }
                    if (control is RichTextBox)
                    {
                        (control as RichTextBox).Clear();
                    }
                    else
                    {
                        func(control.Controls);
                    }
                    if (control is CheckBox)
                    {
                        (control as CheckBox).Checked = true;
                    }
                    else
                    {
                        func(control.Controls);
                    }
                }

            };

            func(Controls);
        }

        private void dgv_list_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txt_dayswoked_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_basic_pay_TextChanged(object sender, EventArgs e)
        {

        }
        void calculate_gross()
        {
            Double gross = gm.toNormalDoubleFormat(txt_basic_pay.Text) + gm.toNormalDoubleFormat(txt_reg_ot_b.Text) + gm.toNormalDoubleFormat(txt_dayoffot_b.Text) + gm.toNormalDoubleFormat(txt_legalhol_ot_b.Text) + gm.toNormalDoubleFormat(txt_specialhol_ot_b.Text) + gm.toNormalDoubleFormat(txt_legalhol_pay_b.Text) + gm.toNormalDoubleFormat(txt_specialhol_pay_b.Text) + gm.toNormalDoubleFormat(txt_night_diff_b.Text) + gm.toNormalDoubleFormat(txt_other_earning.Text) + gm.toNormalDoubleFormat(txt_vl_b.Text) + gm.toNormalDoubleFormat(txt_sl_b.Text) + gm.toNormalDoubleFormat(txt_pl_b.Text) + gm.toNormalDoubleFormat(txt_legalhol_pay_b.Text) + gm.toNormalDoubleFormat(txt_specialhol_pay_b.Text);
              
            lbl_total_gross.Text = gm.toAccountingFormat(gross); 
            
        }
        void total_deduction() {
            Double deductions = gm.toNormalDoubleFormat(txt_sss_a.Text) + gm.toNormalDoubleFormat(txt_philhealth_a.Text) + gm.toNormalDoubleFormat(txt_pagibig_a.Text) + gm.toNormalDoubleFormat(txt_wtax.Text) + gm.toNormalDoubleFormat(txt_other_deductions.Text) + gm.toNormalDoubleFormat(txt_advance_loans.Text) + gm.toNormalDoubleFormat(txt_others.Text)  ;
            lbl_total_tax.Text = deductions.ToString("0.00");
           
        }
        private void btn_compute_gross_Click(object sender, EventArgs e)
        {
            calculate_gross();
        }
        void calculate_net()
        {
            total_deduction();
            Double net = 0.00;
            if (lbl_total_gross.Text != "0.00" || lbl_total_gross.Text != "" || lbl_total_gross.Text != "0")
            {
                lbl_total_net.Text = gm.toAccountingFormat(gm.toNormalDoubleFormat(lbl_total_gross.Text) - gm.toNormalDoubleFormat(lbl_total_tax.Text));
            }
            else
            {
                MessageBox.Show("Make sure that Gross Amount is not Zero.");
            }
        
        }
        private void btn_compute_net_pay_Click(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void btn_compute_payroll_Click(object sender, EventArgs e)
        {

        }

        private void txt_reg_ot_a_TextChanged(object sender, EventArgs e)
        {
            this.reg_ot_b_total = (hour_rate * 1.25) * gm.toNormalDoubleFormat(txt_reg_ot_a.Text);
            txt_reg_ot_b.Text = this.reg_ot_b_total.ToString("0.00");
            calculate_gross();
            calculate_net();
        }

        private void txt_dayoffot_a_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Double total_dayoff_ot =  hour_rate * Convert.ToDouble(txt_dayoffot_a.Text);
                txt_dayoffot_b.Text = total_dayoff_ot.ToString("0.00");
            }
            catch
            {
                txt_dayoffot_b.Text = "0.00";
            }
            calculate_gross();
            calculate_net();
            
        }

        private void txt_legalhol_ot_a_TextChanged(object sender, EventArgs e)
        {
            
            try
            {
                Double total_legal_ot = hour_rate * 2.00 * Convert.ToDouble(txt_legalhol_ot_a.Text);
                txt_legalhol_ot_b.Text = total_legal_ot.ToString("0.00");
            }
            catch
            {
                txt_legalhol_ot_b.Text = "0.00";
            }
            calculate_gross();
            calculate_net();
        }

        private void txt_specialhol_ot_a_TextChanged(object sender, EventArgs e)
        {
            try 
            {
                Double total_specialhol_ot = (hour_rate * 1.3) * Convert.ToDouble(txt_specialhol_ot_a.Text);
                txt_specialhol_ot_b.Text = total_specialhol_ot.ToString("0.00");
            }
            catch 
            {
                txt_specialhol_ot_b.Text = "0.00";
            }
            calculate_gross();
            calculate_net();
        }

        private void txt_legalhol_pay_a_TextChanged(object sender, EventArgs e)
        {
            Double var = 0.00;
            var =  (hour_rate * 2.60) * gm.toNormalDoubleFormat(txt_legalhol_pay_a.Text);
            if (var <= 0 || var <= 0.00)
            {

                var = 0.00;
            }
            txt_legalhol_pay_b.Text = gm.toAccountingFormat(var);
            calculate_gross();
            calculate_net();
        }

        private void txt_specialhol_pay_a_TextChanged(object sender, EventArgs e)
        {
            Double var = 0.00;
           
            var = (hour_rate * 1.69) * gm.toNormalDoubleFormat(txt_specialhol_pay_a.Text);

            if (var <= 0 || var <= 0.00)
            {
                var = 0.00;
            }
            txt_specialhol_pay_b.Text = gm.toAccountingFormat(var);
            calculate_gross();
            calculate_net();
        }

       

        private void txt_dayoffot_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_legalhol_ot_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_specialhol_ot_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_legalhol_pay_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_specialhol_pay_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_vl_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_sl_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_pl_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_night_diff_b_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_other_earning_TextChanged(object sender, EventArgs e)
        {
            calculate_gross();
        }

        private void txt_others_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void txt_advance_loans_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void txt_other_deductions_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void txt_sss_con_c_TextChanged(object sender, EventArgs e)
        {
           // calculate_net();
        }

        private void txt_sss_b_TextChanged(object sender, EventArgs e)
        {
           // calculate_net();
        }

        private void txt_philhealth_b_TextChanged(object sender, EventArgs e)
        {
            //calculate_net();
        }

        private void txt_pagibig_b_TextChanged(object sender, EventArgs e)
        {
           // calculate_net();
        }

        private void txt_sss_a_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void txt_philhealth_a_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void txt_pagibig_a_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void txt_wtax_TextChanged(object sender, EventArgs e)
        {
            calculate_net();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            int r = -1;
            String code = "", name = "", empid = "", ppid = "";
            try
            {
                if (dgv_list.Rows.Count > 1)
                {
                    r = dgv_list.CurrentRow.Index;

                    try
                    {
                        code = dgv_list["pay_code", r].Value.ToString();
                        display_emp_payroll(code, empid, ppid);
                        goto_win2();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Select employee payroll.");
                    }


                }
                else
                {
                    MessageBox.Show("Payroll list is empty.");
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void txt_reg_ot_a_Leave(object sender, EventArgs e)
        {

        }

        private void txt_absent_TextChanged(object sender, EventArgs e)
        {

            try
            {
                absent_amt = Convert.ToDouble(txt_absent.Text) * this.daily_rate;

                basic_pay = (pay_rate - (late_amt + absent_amt));
                txt_basic_pay.Text = basic_pay.ToString("0.00");
                txt_absent_amount.Text = absent_amt.ToString("0.00");
            }
            catch 
            {
                
            }
            
            calculate_gross();
            calculate_net();
            
        }

        private void txt_vl_a_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Double total_vl = Convert.ToDouble(txt_vl_a.Text) * daily_rate;
                txt_vl_b.Text = total_vl.ToString("0.00");
                
            }
            catch 
            {
                txt_vl_b.Text = "0.00";
            }
            calculate_gross();
            calculate_net();
        }

        private void txt_sl_a_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Double total_sl = Convert.ToDouble(txt_sl_a.Text) * daily_rate;
                txt_sl_b.Text = total_sl.ToString("0.00");
            }
            catch
            {
                txt_sl_b.Text = "0.00";
            }
            calculate_gross();
            calculate_net();
        }

        private void txt_pl_a_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Double total_pl = Convert.ToDouble(txt_pl_a.Text) * daily_rate;
                txt_pl_b.Text = total_pl.ToString("0.00");
            }
            catch
            {
                txt_pl_b.Text = "0.00";
            }
            calculate_gross();
            calculate_net();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            disp_payroll();
        }
    }
}
