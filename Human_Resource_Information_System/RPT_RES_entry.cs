using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft;
using System.Drawing.Printing;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Dynamic;

namespace Human_Resource_Information_System
{
    public partial class RPT_RES_entry : Form
    {

        String rpt_no = "";
        String _schema = "";
        thisDatabase db;
        GlobalMethod gm;
        GlobalClass gc;

        ReportDocument myReportDocument;
        ParameterFieldDefinition crParameterFieldDefinition;
        ParameterValues crParameterValues;
        ParameterDiscreteValue crParameterDiscreteValue;
        ParameterFieldDefinitions crParameterFieldDefinitions;

        String fileloc_acctg = "";
        String fileloc_hotel = "";
        String fileloc_inv = "";
        String fileloc_md = "";
        String fileloc_sale = "";
        String fileloc_proj = "";
        String fileloc_srvc = "";
        String fileloc_payroll = "";
        String fileloc_timeclock = "";
        String comp_name, comp_addr, comp_sss, comp_philhealth, comp_pagibig, comp_tin;
        Boolean ishide_opt = true;

        Boolean hasBranch = false;

        Boolean isReady = false, isMain = true;

        String _fy = "";
        dynamic data = new ExpandoObject();

        public RPT_RES_entry(String rnum)
        {
            InitializeComponent();
            db = new thisDatabase();
            gm = new GlobalMethod();
            gc = new GlobalClass();
            myReportDocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            crParameterValues = new ParameterValues();
            crParameterDiscreteValue = new ParameterDiscreteValue();

            DateTime this_t_date = Convert.ToDateTime(db.get_systemdate(""));
            String system_loc = db.get_system_loc();

            rpt_no = rnum;
            fileloc_proj = system_loc + "\\Reports\\Project\\";

            fileloc_acctg = system_loc + "\\Reports\\Accounting\\";
            fileloc_hotel = system_loc + "\\Reports\\Hotel\\";
            fileloc_inv = system_loc + "\\Reports\\Inventory\\";
            fileloc_md = system_loc + "\\Reports\\MD\\";
            fileloc_srvc = system_loc + "\\Reports\\Service\\";
            fileloc_sale = system_loc + "\\Reports\\Sale\\";
            fileloc_payroll = system_loc + "\\Reports\\Payroll\\";
            fileloc_timeclock = system_loc + "\\Reports\\Timeclock\\";

            
            _fy = db.get_m99fy();
            dtp_frm.Value = this_t_date;
            dtp_to.Value = this_t_date;



            try
            {

            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }

            this.Text = this.Text + " | " + rpt_no;
            isReady = true;
        }
        public RPT_RES_entry(String rnum, String title)
            : this(rnum)
        {
            this.Text = title;
            isMain = false;
        }


        private void RPT_RES_entry_Load(object sender, EventArgs e)
        {
            pbar_panl_hide();
        }

        private void btn_submit_Click(object sender, EventArgs e)
        {
            this.print();
        }
        public void print()
        {
            bgworker.RunWorkerAsync();
        }



        private void btn_clear_Click(object sender, EventArgs e)
        {
            cbo_1.SelectedIndex = -1;
            cbo_2.SelectedIndex = -1;
            cbo_3.SelectedIndex = -1;

            chk_1.Checked = false;
            chk_2.Checked = false;
            chk_3.Checked = false;

            cbo_4.Text = "ALL BRANCH"; /*GlobalClass.branch*/

        }

        private int get_cbo_index(ComboBox cbo)
        {
            int i = -1;

            cbo.Invoke(new Action(() => {
                try { i = cbo.SelectedIndex; } catch { }
            }));

            return i;
        }

        private String get_cbo_value(ComboBox cbo)
        {
            String value = "";
            cbo.Invoke(new Action(() => {
                try { value = cbo.SelectedValue.ToString(); } catch { }
            }));
            return value;
        }

        private String get_cbo_text(ComboBox cbo)
        {
            String txt = "";

            cbo.Invoke(new Action(() => {
                try { txt = cbo.Text.ToString(); } catch { }
            }));

            return txt;
        }
        private void set_cbo_droppeddown(ComboBox cbo, Boolean isdrop)
        {
            cbo.Invoke(new Action(() => {
                cbo.DroppedDown = isdrop;
            }));
        }
        private void reset()
        {
            reset_pbar();
            input_enable(true);
            pbar_panl_hide();
        }
        private Boolean ischkbox_checked(CheckBox chk)
        {
            Boolean ischk = false;

            chk.Invoke(new Action(() => {
                try { ischk = chk.Checked; } catch { }
            }));

            return ischk;
        }
        private void chkbox_check(CheckBox chk, Boolean ischk)
        {
            chk.Invoke(new Action(() => {
                try { chk.Checked = ischk; } catch { }
            }));
        }
        private String chkbox_text(CheckBox chk)
        {
            String txt = "";
            chk.Invoke(new Action(() => {
                try { txt = chk.Text.ToString(); } catch { }
            }));
            return txt;
        }
        private void add_fieldparam(String col, String val)
        {
            try
            {
                crParameterDiscreteValue.Value = val;
                crParameterFieldDefinitions = myReportDocument.DataDefinition.ParameterFields;
                crParameterFieldDefinition = crParameterFieldDefinitions[col];
                crParameterValues = crParameterFieldDefinition.CurrentValues;
                clr_param();
                inc_pbar(10);
            }
            catch (Exception error) { MessageBox.Show(error.Message); }
        }

        private void disp_reportviewer(ReportDocument myReportDocument)
        {
            crptviewer.Invoke(new Action(() => {
                try { crptviewer.ReportSource = myReportDocument; } catch { }
            }));

            crptviewer.Invoke(new Action(() => {
                crptviewer.Refresh();
            }));
        }

        private void clr_param()
        {
            crParameterValues.Clear();
            crParameterValues.Add(crParameterDiscreteValue);
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);
        }

        private void input_enable(Boolean bol)
        {
            if (!isMain) return;

            cbo_1.Invoke(new Action(() => {
                try { cbo_1.Enabled = bol; } catch { }
            }));

             cbo_2.Invoke(new Action(() => {
                try { cbo_2.Enabled = bol; } catch { }
            }));

            cbo_3.Invoke(new Action(() => {
                try { cbo_3.Enabled = bol; } catch { }
            }));

            dtp_frm.Invoke(new Action(() => {
                try { dtp_frm.Enabled = bol; } catch { }
            }));

            dtp_to.Invoke(new Action(() => {
                try { dtp_to.Enabled = bol; } catch { }
            }));

            btn_clear.Invoke(new Action(() => {
                try { btn_clear.Enabled = bol; } catch { }
            }));

            btn_submit.Invoke(new Action(() => {
                try { btn_submit.Enabled = bol; } catch { }
            }));

            cbo_4.Invoke(new Action(() => {
                try { cbo_4.Enabled = bol; } catch { }
            }));

        }

        private void inc_pbar(int i)
        {
            try
            {
                progressBar1.Invoke(new Action(() => {
                    progressBar1.Value += i;
                }));
            }
            catch (Exception) { reset_pbar(); }
        }

        private void reset_pbar()
        {
            progressBar1.Invoke(new Action(() => {
                try { progressBar1.Value = 0; } catch { }
            }));
        }

        private void pbar_panl_hide()
        {
            pnl_pbar.Invoke(new Action(() => {
                pnl_pbar.Hide();
            }));
        }

        private void pbar_panl_show()
        {
            pnl_pbar.Invoke(new Action(() => {
                pnl_pbar.Show();
            }));
        }

        private void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            input_enable(false);

            String dtpicker_frm = dtp_frm.Value.ToString("MM-dd-yyyy");
            String dtpicker_to = dtp_to.Value.ToString("MM-dd-yyyy");
            String branch_cbo = "";
            String WHERE = "";

            _schema = db.get_schema();
            comp_name = db.get_m99comp_name();
            comp_addr = db.get_m99comp_addr();
            comp_sss = db.get_colval("m99", "comp_sss", "");
            comp_philhealth = db.get_colval("m99", "comp_philhealth", "");
            comp_pagibig = db.get_colval("m99", "comp_pagibig", "");
            comp_tin = db.get_colval("m99", "comp_tin", "");

            pbar_panl_show();

            try
            {
                if (rpt_no == "M001")
                {
                    String cat = "";
                    String brd = "";

                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "employee_listing.rpt");

                    DataTable dt = null;

                    dt = db.QueryOnTableWithParams("items", "*", WHERE, " ORDER BY item_desc ASC");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }

                else if (rpt_no == "P101")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "payroll_summary_report.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("report_type", " - " + data.report_type);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P102")
                {
                    inc_pbar(10);
                   
                    myReportDocument.Load(fileloc_payroll + "payslip.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P201")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "whtax_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("comp_tin", comp_tin);
                    add_fieldparam("fymo", data.fymo);

                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P202")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "sss_contr_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("comp_sss", comp_sss);
                    add_fieldparam("fymo", data.fymo);
                    
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P203")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "philhealth_contr_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("comp_philhealth", comp_philhealth);
                    add_fieldparam("fymo", data.fymo);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P204")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "pagibig_contr_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("comp_pagibig", comp_pagibig);
                    add_fieldparam("fymo", data.fymo);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P205")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "13th_monthpay_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("fy", data.fy);
                    add_fieldparam("mo", data.mo);
                    add_fieldparam("empstatus", data.empstatus);

                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P206")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "gross_pay_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("fy", data.fy);
                    add_fieldparam("mo", data.mo);
                    add_fieldparam("empstatus", data.empstatus);

                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P207")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "leave_status_report.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("empt_name", data.empt_name);
                    add_fieldparam("fy", data.fy);
                    
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P208")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "cumulative_payroll.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("fy", data.fy);
                    add_fieldparam("mo", data.mo);

                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "P209")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "leave_balance_report.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);

                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "E103")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "employee_listing.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("dept_name", data.dept_name);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "M101")
                {
                    inc_pbar(10);
                    String file_name = "";
                    if (data.hasEmp)
                    {
                        file_name = "employee_info.rpt";
                    }
                    else
                    {
                        file_name = "employee_list.rpt";
                    }

                    myReportDocument.Load(fileloc_payroll + file_name);

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);


                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }

                else if (rpt_no == "M102")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "other_earning_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);


                    add_fieldparam("earning", data.earning);
                    add_fieldparam("pay_period", data.pay_period);
                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "M103")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "other_deduction_summary.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("deduction", data.deduction);
                    add_fieldparam("pay_period", data.pay_period);
                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "M104")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "other_deduction_per_employee.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("deduction", data.deduction);
                    add_fieldparam("pay_period", data.pay_period);
                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "M105")
                {
                    inc_pbar(10);
                    String filename = "employee_loan_summary.rpt";
                    if (!String.IsNullOrEmpty(data.t_date))
                    {
                        filename = "employee_loan_summary_transdate.rpt";
                    
                    }
                    myReportDocument.Load(fileloc_payroll + filename);

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("t_date", data.t_date);
                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "M105e")
                {
                    inc_pbar(10);

                    myReportDocument.Load(fileloc_payroll + "employee_loan.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD01")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_department.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT deptid AS department, dept_name  FROM rssys.hr_department WHERE COALESCE(cancel,'')<>'Y'");
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD02")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_section.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT secid AS section, section_name, dept_name FROM rssys.hr_depsection s LEFT JOIN rssys.hr_department d ON d.deptid=s.deptid WHERE COALESCE(s.cancel,'')<>'Y'");
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD03")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_jobtitle.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT postid AS positions, position_name AS jtitle_name FROM rssys.hr_position WHERE COALESCE(cancel,'')<>'Y'");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD04")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_shift_schedule.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS shift_sched, remarks AS shift_sched_remarks, type AS shift_sched_type, CASE WHEN ss.am_timein<>'' AND ss.pm_timein<>'' THEN ss.am_timein ||' - '|| ss.am_timeout ||'AM/'|| ss.pm_timein ||' - '|| ss.pm_timeout||'PM' WHEN ss.am_timein<>'' THEN ss.am_timein ||' - '|| ss.am_timeout||'AM' ELSE ss.pm_timein ||' - '|| ss.pm_timeout||'PM' END AS shift_sched_desc FROM rssys.hr_shift_schedule ss WHERE COALESCE(cancel,'')<>'Y'");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD05")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_holiday.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT to_char(date_holiday,'YYYY-MM-DD') AS id1, description AS desc1, CASE WHEN holiday_type='L' THEN 'Legal' ELSE 'Special' END AS desc2 FROM rssys.hr_holidays WHERE COALESCE(cancel,'')<>'Y' ORDER BY date_holiday");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD06")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_payroll_period.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT pay_code AS id1,to_char(date_from,'YYYY/MM/DD')||' - '||to_char(date_to,'YYYY/MM/DD') AS name1,CASE WHEN d_w_tax='N' THEN 'NO' ELSE 'YES' END AS desc5,CASE WHEN d_sss_c='N' THEN 'NO' ELSE 'YES' END AS desc6,CASE WHEN d_philhealth='N' THEN 'NO' ELSE 'YES' END AS desc7,CASE WHEN d_pagibig='N' THEN 'NO' ELSE 'YES' END AS desc8,financial_year AS desc1,month AS desc2,pay_type AS desc3,num_days AS desc4,payroll_classic AS desc9 FROM rssys.hr_payrollpariod WHERE COALESCE(cancel,'')<>'Y' ORDER BY pay_code");
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD07")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_withholding_tax.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, description AS desc1, exemption AS amount0,bracket1 AS amount1,bracket2 AS amount2,bracket3 AS amount3,bracket4 AS amount4,bracket5 AS amount5,bracket6 AS amount6,bracket7 AS amount7,bracket8 AS amount8,bracket9 AS amount9,bracket10 AS amount10 FROM rssys.hr_wtax WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD08")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_sss_bracket.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, bracket1 AS amount1, bracket2 AS amount2, s_credit AS amount3, empshare_sc AS amount4, s_ec AS amount5, empshare_ec AS amount6 FROM rssys.hr_sss WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD09")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_philhealth_bracket.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1,bracket1 AS amount1,bracket2 AS amount2,salary_base AS amount3,emp_er AS amount4, emp_ee AS amount5 FROM rssys.hr_philhealth WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD10")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_pagibig_bracket.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, bracket1 AS amount1, bracket2 AS amount2, pct AS amount3 FROM rssys.hr_hdmf WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD11")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_loan.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, description AS name1 FROM rssys.hr_loan_type WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");
                    
                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD12")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_otherearning.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, description AS name1 FROM rssys.hr_other_earnings WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD13")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_otherdeduction.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, description AS name1 FROM rssys.hr_other_deductions WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "MD14")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_payroll + "md_leaves.rpt");

                    DataTable dt = db.QueryBySQLCode("SELECT code AS id1, description AS name1 FROM rssys.hr_leave_type WHERE COALESCE(cancel,'')<>'Y' ORDER BY code");

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("comp_name", comp_name);
                    add_fieldparam("comp_addr", comp_addr);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "T101")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_timeclock + "employee_dtr_report.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("t_date", data.date);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "T102")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_timeclock + "daily_timelog_report.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("t_date", data.date);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                else if (rpt_no == "T103")
                {
                    inc_pbar(10);
                    myReportDocument.Load(fileloc_timeclock + "absencelateundertime_report.rpt");

                    DataTable dt = db.QueryBySQLCode(data.strSQL);

                    myReportDocument.Database.Tables[0].SetDataSource(dt);
                    inc_pbar(10);

                    add_fieldparam("t_date", data.date);
                    add_fieldparam("userid", GlobalClass.username);

                    disp_reportviewer(myReportDocument);
                }
                
            }
            catch (Exception er)
            {
                throw er;
            }

            reset();
        }

        private void RPT_RES_entry_FormClosing(object sender, FormClosingEventArgs e)
        {
            //bgworker.CancelAsync();
        }

        private void cbo_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isReady)
            {

            }
        }

        private void crptviewer_Load(object sender, EventArgs e)
        {

        }

        private void btn_minimize_Click(object sender, EventArgs e)
        {
            if (ishide_opt)
            {
                pnl_rpt_option.Hide();
                pnl_rpt_option_header.Height = 30;
                btn_minimize.Width = 150;

                ishide_opt = false;
            }
            else
            {
                pnl_rpt_option.Show();
                pnl_rpt_option_header.Height = 114;
                btn_minimize.Width = 64;

                ishide_opt = true;
            }
        }

        private void grp_options_Enter(object sender, EventArgs e)
        {

        }

        private void dtp_frm_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dtp_to_ValueChanged(object sender, EventArgs e)
        {

        }

        private void cbo_2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isReady)
            {

            }
        }

        private void cbo_3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isReady)
            {

            }
        }

        private void chk_1_CheckedChanged(object sender, EventArgs e)
        {
            if (isReady)
            {
                if (ischkbox_checked(chk_1))
                {
                }
            }
        }
        private void chk_2_CheckedChanged(object sender, EventArgs e)
        {
            if (isReady)
            {
                if (ischkbox_checked(chk_2))
                {

                }
            }
        }

        private void chk_3_CheckedChanged(object sender, EventArgs e)
        {
            if (isReady)
            {
                if (ischkbox_checked(chk_3))
                {

                }
            }
        }

        private void cbo_4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        public void print_master_data()
        {
            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_employees(String empid)
        {
            String strSQL = "";

            if (!String.IsNullOrEmpty(empid))
            {
                strSQL = "SELECT d.dept_name, ds.section_name, j.jtitle_name, e.empid, e.lastname,e.firstname,e.mi,e.section,e.positions,e.sex,CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.birth,'YYYY-MM-DD') END AS birth,e.civil_status,e.home_address, e.home_tel, e.empstatus,CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_hired,'YYYY-MM-DD') END AS date_hired, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_regular,'YYYY-MM-DD') END AS date_regular, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_terminated,'YYYY-MM-DD') END AS date_terminated, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_resigned,'YYYY-MM-DD') END AS date_resigned, e.sss, e.philhealth, e.pagibig, e.tin, e.tax_bracket, e.shift_sched,e.rate_type, e.pay_rate, e.dayoff1, e.dayoff2, CASE WHEN ss.am_timein<>'' AND ss.pm_timein<>'' THEN ss.am_timein ||' - '|| ss.am_timeout ||'AM/'|| ss.pm_timein ||' - '|| ss.pm_timeout||'PM' WHEN ss.am_timein<>'' THEN ss.am_timein ||' - '|| ss.am_timeout||'AM' ELSE ss.pm_timein ||' - '|| ss.pm_timeout||'PM' END AS shift_sched_desc FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON d.deptid=e.department LEFT JOIN rssys.hr_depsection ds ON ds.secid=e.section LEFT JOIN rssys.hr_jobtitle j ON j.jtid=e.positions LEFT JOIN rssys.hr_shift_schedule ss ON ss.code=e.shift_sched WHERE e.empid='" + empid + "'";
            }
            else
            {
                strSQL = "SELECT d.dept_name, ds.section_name, j.jtitle_name, e.empid, e.lastname,e.firstname,e.mi,e.section,e.positions,e.sex, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.birth,'YYYY-MM-DD') END AS birth,e.civil_status,e.home_address, e.home_tel, e.empstatus,CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_hired,'YYYY-MM-DD') END AS date_hired, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_regular,'YYYY-MM-DD') END AS date_regular, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_terminated,'YYYY-MM-DD') END AS date_terminated, CASE WHEN e.date_terminated='1900-01-01' THEN '' ELSE to_char(e.date_resigned,'YYYY-MM-DD') END AS date_resigned,e.sss, e.philhealth, e.pagibig, e.tin, e.tax_bracket, e.shift_sched,e.rate_type, e.pay_rate, e.dayoff1, e.dayoff2, CASE WHEN ss.am_timein<>'' AND ss.pm_timein<>'' THEN ss.am_timein ||' - '|| ss.am_timeout ||'AM/'|| ss.pm_timein ||' - '|| ss.pm_timeout||'PM' WHEN ss.am_timein<>'' THEN ss.am_timein ||' - '|| ss.am_timeout||'AM' ELSE ss.pm_timein ||' - '|| ss.pm_timeout||'PM' END AS shift_sched_desc FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON d.deptid=e.department LEFT JOIN rssys.hr_depsection ds ON ds.secid=e.section LEFT JOIN rssys.hr_jobtitle j ON j.jtid=e.positions LEFT JOIN rssys.hr_shift_schedule ss ON ss.code=e.shift_sched";
            }

            data.strSQL = strSQL;
            data.hasEmp = !String.IsNullOrEmpty(empid);

            pnl_rpt_option_header.Hide();
            print();
        }

        public void print_other_earning_summary(String payperiod_code, String payperiod_desc, String earncode_frm, String earntxt_frm, String earncode_to, String earntxt_to)
        {
            String WHERE = "", strSQL = "";
            String earn_desc = "ALL";
            
            if (!String.IsNullOrEmpty(earncode_frm))
            {
                WHERE = " AND '" + earncode_frm + "' <= ee.earning_code ";
                earn_desc = earntxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(earncode_to))
            {
                WHERE = " AND '" + earncode_to + "' >= ee.earning_code ";
                earn_desc = "First To " + earntxt_to;
            }
            if (!String.IsNullOrEmpty(earncode_frm) && !String.IsNullOrEmpty(earncode_to))
            {
                WHERE = " AND ee.earning_code BETWEEN '" + earncode_frm + "'  AND '" + earncode_to + "' ";
                if (earncode_frm == earncode_to) earn_desc = earntxt_frm;
                else  earn_desc = earntxt_frm + " To " + earntxt_to;
               
            }
            WHERE += " AND payroll_period='" + payperiod_code + "'";

            strSQL = "SELECT ee.*, oe.description AS earn_desc FROM rssys.hr_earning_entry ee LEFT JOIN rssys.hr_other_earnings oe ON oe.code= ee.earning_code WHERE 1=1 " + WHERE + "";

            data.strSQL = strSQL;
            data.earning = earn_desc;
            data.pay_period = payperiod_desc;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_philhealth_contribution_summary(String fy, String mofrm, String moto, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to)
        { 
            String WHERE = "", WHERE2 = "", strSQL = "";
            String dept_name = "";
            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE2 = " AND '" + deptcode_frm + "' <= e.department ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND '" + deptcode_to + "' >= e.department ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND e.department BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }


            WHERE = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT COALESCE(s.emp_ee,0) AS amount1, COALESCE(s.emp_er,0) AS amount2, (COALESCE(s.emp_ee,0) + COALESCE(s.emp_er,0)) AS amount3, e.empid AS id1, e.lastname,e.firstname, e.mi, (e.lastname||', '||e.firstname||', '||e.mi) AS name1, e.sss, e.pagibig, e.philhealth AS id2, to_char((pp.financial_year||'-'|| LPAD(pp.month,2,'0') ||'-01')::date,'YYYY MON') AS name2, pp.financial_year, pp.month, TO_CHAR(pp.date_from,'YYYY-MM-DD') AS date_from,  TO_CHAR(pp.date_to,'YYYY-MM-DD') AS date_to FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_payrollpariod pp ON (ep.ppid=pp.pay_code  " + WHERE + ") LEFT JOIN rssys.hr_employee e ON (e.empid=ep.empid " + WHERE2 + ") LEFT JOIN rssys.hr_philhealth s ON (s.code=ep.philhealth_bracket) ORDER BY e.empid, pp.financial_year, pp.month";

            data.strSQL = strSQL;
            data.fymo = DateTime.Parse(fy + "-" + mofrm + "-01").ToString("MMMM") + " - " + DateTime.Parse(fy + "-" + moto + "-01").ToString("MMMM") + " " + fy;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_pagibig_contribution_summary(String fy, String mofrm, String moto, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to)
        {
            String WHERE = "", WHERE2 = "", strSQL = "";
            String dept_name = "";
            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE2 = " AND '" + deptcode_frm + "' <= e.department ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND '" + deptcode_to + "' >= e.department ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND e.department BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }


            WHERE = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT COALESCE(s.emp_ee,0) AS amount1, COALESCE(s.emp_er,0) AS amount2, (COALESCE(s.emp_ee,0) + COALESCE(s.emp_er,0)) AS amount3, e.empid AS id1, e.lastname,e.firstname, e.mi, (e.lastname||', '||e.firstname||', '||e.mi) AS name1, e.sss, e.pagibig AS id2, e.philhealth, to_char((pp.financial_year||'-'|| LPAD(pp.month,2,'0') ||'-01')::date,'YYYY MON') AS name2, pp.financial_year, pp.month, TO_CHAR(pp.date_from,'YYYY-MM-DD') AS date_from,  TO_CHAR(pp.date_to,'YYYY-MM-DD') AS date_to FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_payrollpariod pp ON (ep.ppid=pp.pay_code  " + WHERE + ") LEFT JOIN rssys.hr_employee e ON (e.empid=ep.empid " + WHERE2 + ") LEFT JOIN rssys.hr_hdmf s ON (s.code=ep.pagibig_bracket) ORDER BY e.empid, pp.financial_year, pp.month";

            data.strSQL = strSQL;
            data.fymo = DateTime.Parse(fy + "-" + mofrm + "-01").ToString("MMMM") + " - " + DateTime.Parse(fy + "-" + moto + "-01").ToString("MMMM") + " " + fy;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_sss_contribution_summary(String fy, String mofrm, String moto, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to)
        {
            String WHERE = "", WHERE2 = "", strSQL = "";
            String dept_name = "";
            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE2 = " AND '" + deptcode_frm + "' <= e.department ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND '" + deptcode_to + "' >= e.department ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND e.department BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }



            WHERE = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT COALESCE(s.empshare_sc,0) AS amount1, COALESCE(s.empshare_ec,0) AS amount2, COALESCE(s.s_ec,0) AS amount3, (COALESCE(s.empshare_sc,0) + COALESCE(s.s_ec,0) + COALESCE(s.empshare_ec,0)) AS amount4, e.empid AS id1, e.lastname,e.firstname, e.mi, (e.lastname||', '||e.firstname||', '||e.mi) AS name1, e.sss AS id2, e.pagibig, e.philhealth, to_char((pp.financial_year||'-'|| LPAD(pp.month,2,'0') ||'-01')::date,'YYYY MON') AS name2, pp.financial_year, pp.month, TO_CHAR(pp.date_from,'YYYY-MM-DD') AS date_from,  TO_CHAR(pp.date_to,'YYYY-MM-DD') AS date_to FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_payrollpariod pp ON (ep.ppid=pp.pay_code  " + WHERE + ") LEFT JOIN rssys.hr_employee e ON (e.empid=ep.empid " + WHERE2 + ") LEFT JOIN rssys.hr_sss s ON (s.code=ep.sss_bracket)  ORDER BY e.empid, pp.financial_year, pp.month";

            data.strSQL = strSQL;
            data.fymo = DateTime.Parse(fy + "-" + mofrm + "-01").ToString("MMMM") + " - " + DateTime.Parse(fy + "-" + moto + "-01").ToString("MMMM") + " " + fy;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_witholdingtax_summary(String fy, String mofrm, String moto, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to)
        {
            String WHERE = "", WHERE2 = "", strSQL = "";
            String dept_name = "";
            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE2 = " AND '" + deptcode_frm + "' <= e.department ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND '" + deptcode_to + "' >= e.department ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE2 = " AND e.department BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }



            WHERE = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT ep.w_tax AS amount1, e.empid AS id1, e.lastname,e.firstname, e.mi, (e.lastname||', '||e.firstname||', '||e.mi) AS name1, e.tin AS id2, e.pagibig, e.philhealth, to_char((pp.financial_year||'-'|| LPAD(pp.month,2,'0') ||'-01')::date,'YYYY MON') AS name2, pp.financial_year, pp.month, TO_CHAR(pp.date_from,'YYYY-MM-DD') AS date_from,  TO_CHAR(pp.date_to,'YYYY-MM-DD') AS date_to FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_payrollpariod pp ON (ep.ppid=pp.pay_code  " + WHERE + ") LEFT JOIN rssys.hr_employee e ON (e.empid=ep.empid " + WHERE2 + ") ORDER BY e.empid, pp.financial_year, pp.month";

            data.strSQL = strSQL;
            data.fymo = DateTime.Parse(fy + "-" + mofrm + "-01").ToString("MMMM") + " - " + DateTime.Parse(fy + "-" + moto + "-01").ToString("MMMM") + " " + fy; 

            pnl_rpt_option_header.Hide();
            print();
        }


        public void print_13monthpay_summary(String fy, String mofrm, String moto, String empstatus_text, String empstatus_code)
        {

            String WHERE = "", WHERE2 = "", strSQL = "";
            String empstatus = "All";

            if (!String.IsNullOrEmpty(empstatus_code))
            {
                WHERE = " AND e.empstatus='" + empstatus_code + "'";
                empstatus = empstatus_text;
            }

            WHERE2 = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT ep.*, d.dept_name, pp.pay_code, pp.financial_year AS fy ,pp.month AS mo, to_char(pp.date_from,'YYYY-MM-DD') AS date_from, to_char(pp.date_to,'YYYY-MM-DD') AS date_to, to_char(pp.date_from,'Month DD, YYYY') AS date_from_desc, to_char(pp.date_to,'Month DD, YYYY') AS date_to_desc, e.firstname, e.lastname, e.mi, e.tax_bracket FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee e ON (e.empid=ep.empid " + WHERE + ") LEFT JOIN rssys.hr_payrollpariod pp ON (pp.pay_code=ep.ppid " + WHERE2 + ") LEFT JOIN rssys.hr_department d ON e.department=d.deptid ORDER BY e.empid, pp.date_from DESC";
             
            data.strSQL = strSQL;
            data.empstatus = empstatus;
            data.fy = fy;
            data.mo = mofrm + " - " + moto;

            pnl_rpt_option_header.Hide();
            print();
        }

        public void print_grosspay_summary(String fy, String mofrm, String moto, String empstatus_text, String empstatus_code)
        {

            String WHERE = "", WHERE2 = "", strSQL = "";
            String empstatus = "All";

            if (!String.IsNullOrEmpty(empstatus_code))
            {
                WHERE = " AND e.empstatus='" + empstatus_code + "'";
                empstatus = empstatus_text;
            }

            WHERE2 = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT ep.*, d.dept_name, pp.pay_code, pp.financial_year AS fy ,pp.month AS mo, to_char(pp.date_from,'YYYY-MM-DD') AS date_from, to_char(pp.date_to,'YYYY-MM-DD') AS date_to, to_char(pp.date_from,'Month DD, YYYY') AS date_from_desc, to_char(pp.date_to,'Month DD, YYYY') AS date_to_desc, e.firstname, e.lastname, e.mi, e.tax_bracket FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee e ON (e.empid=ep.empid " + WHERE + ") LEFT JOIN rssys.hr_payrollpariod pp ON (pp.pay_code=ep.ppid " + WHERE2 + ") LEFT JOIN rssys.hr_department d ON e.department=d.deptid ORDER BY e.empid, pp.date_from DESC";

            data.strSQL = strSQL;
            data.empstatus = empstatus;
            data.fy = fy;
            data.mo = mofrm + " - " + moto;

            pnl_rpt_option_header.Hide();
            print();
        }




        public void print_other_deduction_summary(String payperiod_code, String payperiod_desc, String dedcode_frm, String dedtxt_frm, String dedcode_to, String dedtxt_to)
        {
            String WHERE = "", strSQL = "";
            String deduc_desc = "ALL";

            if (!String.IsNullOrEmpty(dedcode_frm))
            {
                WHERE = " AND '" + dedcode_frm + "' <= de.deduction_code ";
                deduc_desc = dedtxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(dedcode_to))
            {
                WHERE = " AND '" + dedcode_to + "' >= de.deduction_code ";
                deduc_desc = "First To " + dedtxt_to;
            }
            if (!String.IsNullOrEmpty(dedcode_frm) && !String.IsNullOrEmpty(dedcode_to))
            {
                WHERE = " AND de.deduction_code BETWEEN '" + dedcode_frm + "'  AND '" + dedcode_to + "' ";
                if (dedcode_frm == dedcode_to) deduc_desc = dedtxt_frm;
                else deduc_desc = dedtxt_frm + " To " + dedtxt_to;
            }
            WHERE += " AND payroll_period='" + payperiod_code + "'";

            strSQL = "SELECT de.*, od.description AS deduct_desc FROM rssys.hr_deduction_entry de LEFT JOIN rssys.hr_other_deductions od ON od.code=de.deduction_code  WHERE 1=1 " + WHERE + "";


            data.strSQL = strSQL;
            data.deduction = deduc_desc;
            data.pay_period = payperiod_desc;

            pnl_rpt_option_header.Hide();
            print();
        }


        public void print_employee_loan_summary(String loancode)
        {
            String WHERE = "", strSQL = "";
           
            strSQL = "SELECT l.*, lt.description AS loan_type_desc FROM rssys.hr_loanhdr l LEFT JOIN rssys.hr_loan_type lt ON (lt.code=l.loan_type) WHERE loan_code='" + loancode + "' ";

            data.strSQL = strSQL;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_employee_loan_summary(String employee, String t_date_frm, String t_date_to)
        {
            String WHERE = "", strSQL = "";
            String t_date = "";

            if (!String.IsNullOrEmpty(employee))
            {
                WHERE += " AND employee_no='" + employee + "'";
            }

            if (!String.IsNullOrEmpty(t_date_frm) && !String.IsNullOrEmpty(t_date_to))
            {
                WHERE += " AND loan_transdate BETWEEN '" + t_date_frm + "' AND '" + t_date_to + "'";
                if (t_date_frm == t_date_to)
                {
                    t_date = t_date_frm;
                }
                else
                {
                    t_date = t_date_frm + " to " + t_date_to;
                }
            }

            strSQL = "SELECT l.*, lt.description AS loan_type_desc FROM rssys.hr_loanhdr l LEFT JOIN rssys.hr_loan_type lt ON (lt.code=l.loan_type) WHERE 1=1 " + WHERE + " ORDER BY employee_no, loan_transdate";


            data.strSQL = strSQL;
            data.t_date = t_date;

            pnl_rpt_option_header.Hide();
            print();
        }


        public void print_other_deduction_summary_per_employee(String payperiodcode_frm, String payperioddesc_frm, String payperiodcode_to, String payperioddesc_to, String dedcode, String dedtxt, String empid)
        {
            String WHERE = "", strSQL = "";
            String deduc_desc = "ALL", payperiod_desc = "ALL";

            if (!String.IsNullOrEmpty(payperiodcode_frm))
            {
                WHERE = " AND '" + payperiodcode_frm + "' <= payroll_period ";
                payperiod_desc = payperioddesc_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(payperiodcode_to))
            {
                WHERE = " AND '" + payperiodcode_to + "' >= payroll_period ";
                payperiod_desc = "First To " + payperioddesc_to;
            }
            if (!String.IsNullOrEmpty(payperiodcode_frm) && !String.IsNullOrEmpty(payperiodcode_to))
            {
                WHERE = " AND payroll_period BETWEEN '" + payperiodcode_frm + "'  AND '" + payperiodcode_to + "' ";
                if (payperiodcode_frm == payperiodcode_to) payperiod_desc = payperioddesc_frm;
                else payperiod_desc = payperioddesc_frm + " To " + payperioddesc_to;
            }

            if (!String.IsNullOrEmpty(dedcode))
            {
                WHERE += " AND de.deduction_code='" + dedcode + "'";
                deduc_desc = dedtxt;
            }
            if (!String.IsNullOrEmpty(empid))
            {
                WHERE += " AND de.emp_no='" + empid + "'";
            }

            strSQL = "SELECT de.*, od.description AS deduct_desc FROM rssys.hr_deduction_entry de LEFT JOIN rssys.hr_other_deductions od ON od.code=de.deduction_code  WHERE 1=1 " + WHERE + "";

            data.strSQL = strSQL;
            data.deduction = deduc_desc;
            data.pay_period = payperiod_desc;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_employee_listing(String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to, Boolean includeEmpRateHistory, Boolean confidentialOnly)
        {
            String WHERE = "", strSQL = "";
            String dept_name = "";
            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }

            strSQL = "SELECT e.empid, e.firstname, e.lastname, e.mi, (e.pay_rate)::text, to_char(e.date_hired,'YYYY-MM-DD') AS date_hired, to_char(e.date_regular,'YYYY-MM-DD') AS date_regular, e.department, d.dept_name, p.position_name, r.description AS rate_type_desc, ephis.rate_types, ephis.pay_rates FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON e.department=d.deptid LEFT JOIN rssys.hr_position p ON e.positions=p.postid LEFT JOIN rssys.hr_rate_type r ON e.rate_type=r.ratecode LEFT JOIN (SELECT DISTINCT empid, string_agg(to_char(payrate_dt,'YYYY-MM-DD')||' ('||rate_type||')','\n') AS rate_types, string_agg((pay_rate)::text,'\n') AS pay_rates FROM rssys.hr_emp_payrate_history GROUP BY empid) ephis ON (ephis.empid=e.empid " + (!includeEmpRateHistory ? " AND 1!=1" : "") + ")  WHERE 1=1 " + WHERE + " ORDER BY empid ASC ";


            data.strSQL = strSQL;
            data.dept_name = dept_name;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_leave_status(String fy, String emptxt_frm, String empcode_frm, String emptxt_to, String empcode_to)
        {

            String WHERE = "", WHERE2 = "", strSQL = "";
            String empt_name = "All";

            if (!String.IsNullOrEmpty(empcode_frm))
            {
                WHERE = " AND '" + empcode_frm + "' <= e.empid ";
                empt_name = emptxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(empcode_to))
            {
                WHERE = " AND '" + empcode_to + "' >= e.empid ";
                empt_name = "First To " + emptxt_to;
            }
            if (!String.IsNullOrEmpty(empcode_frm) && !String.IsNullOrEmpty(empcode_to))
            {
                WHERE = " AND e.empid BETWEEN '" + empcode_frm + "'  AND '" + empcode_to + "' ";
                if (empcode_frm == empcode_to) empt_name = emptxt_frm;
                else empt_name = emptxt_frm + " To " + emptxt_to;
            }

            WHERE = " AND pp.financial_year='" + fy + "' ";

            strSQL = "SELECT ep.leave_type AS name2, e.empid AS id1, e.lastname||', '||e.firstname||' '||e.mi AS name1, SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Jan' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount1,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Feb' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount2,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Mar' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount3,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Apr' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount4,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='May' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount5,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Jun' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount6,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Jul' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount7,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Aug' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount8,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Sep' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount9,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Oct' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount10,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Nov' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount11,SUM(CASE WHEN trim(to_char(pp.date_from,'Mon'))='Dec' THEN COALESCE(cast(ep.hrs as float),0) ELSE 0 END) AS amount12,SUM(ep.hrs) AS amount13, SUM(ep.hrs)*-1 AS amount14 FROM rssys.hr_employee e LEFT JOIN (SELECT ep.ppid, emp_pay_code, empid, COALESCE(cast(sl_a as float),0) AS hrs, COALESCE(cast(sl_b as float),0) AS amount, 'SL' AS leave_type  FROM rssys.hr_emp_payroll ep UNION ALL SELECT ep.ppid, emp_pay_code, empid, COALESCE(cast(pl_a as float),0), COALESCE(cast(pl_b as float),0), 'PL' AS leave_type FROM rssys.hr_emp_payroll ep UNION ALL SELECT ep.ppid, emp_pay_code, empid, COALESCE(cast(vl_a as float),0), COALESCE(cast(vl_b as float),0), 'VL' AS leave_type FROM rssys.hr_emp_payroll ep UNION ALL SELECT ep.ppid, emp_pay_code, empid, COALESCE(cast(ol_a as float),0), COALESCE(cast(leave_days as float),0), 'OL' AS leave_type FROM rssys.hr_emp_payroll ep ORDER BY emp_pay_code) ep ON (ep.empid=e.empid) LEFT JOIN rssys.hr_payrollpariod pp ON pp.pay_code=ep.ppid WHERE COALESCE(ep.leave_type,'')<>'' " + WHERE + " GROUP BY ep.leave_type, e.rate_type, e.pay_rate, e.empid, e.lastname, e.firstname, e.mi ORDER BY e.empid";

            data.strSQL = strSQL;
            data.empt_name = empt_name;
            data.fy = fy;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_leave_balance(String fy, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to)
        {

            String WHERE = "",  strSQL = "";
            String dept_name = "All";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= e.department ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= e.department ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND e.department BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }

            WHERE = " AND pp.financial_year='" + fy + "' ";

            strSQL = "SELECT round(pay_rate / (CASE WHEN rate_type='M' THEN 30 WHEN rate_type='W' THEN 7 WHEN rate_type='D' THEN 1 ELSE 1 END),2) AS pay_rate, e.empid, e.lastname, e.firstname, e.mi, SUM(COALESCE(cast(vl_a as float),0)) AS vl_a, SUM(COALESCE(cast(vl_b as float),0)) AS vl_b, SUM(COALESCE(cast(sl_a as float),0)) AS sl_a, SUM(COALESCE(cast(sl_b as float),0)) AS sl_b, SUM(COALESCE(cast(pl_a as float),0)) AS pl_a, SUM(COALESCE(cast(pl_b as float),0)) AS pl_b, SUM(COALESCE(cast(ol_a as float),0)) AS ol_a, SUM(COALESCE(cast(ol_b as float),0)) AS ol_b, SUM(COALESCE(cast(leave_amnt as float),0)) AS leave_amnt, SUM(COALESCE(cast(leave_days as float),0)) AS leave_days  FROM rssys.hr_employee e LEFT JOIN rssys.hr_emp_payroll ep ON (ep.empid=e.empid) LEFT JOIN rssys.hr_payrollpariod pp ON pp.pay_code=ep.ppid WHERE 1=1 " + WHERE + " GROUP BY e.rate_type, e.pay_rate, e.empid, e.lastname, e.firstname, e.mi";

            data.strSQL = strSQL;
            data.dept_name = dept_name;
            data.fy = fy;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_cummulative_payroll(String fy, String mofrm, String moto, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to)
        {

            String WHERE = "", WHERE2 = "", strSQL = "";
            String dept_name = "All";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }

            WHERE2 = " AND pp.financial_year='" + fy + "' AND pp.month BETWEEN '" + mofrm + "' AND '" + moto + "'";

            strSQL = "SELECT SUM(ep.leave_amnt) AS leave_amnt, SUM(ep.late_amnt) AS late_amnt, SUM(ep.absent_amnt) AS absent_amnt, SUM(ep.basic_pay) AS basic_pay, SUM(cast(ep.reqular_ot_b as float)) AS reqular_ot_b, SUM(cast(ep.dayoff_ot_b as float)) AS dayoff_ot_b, SUM(cast(ep.legal_hol_ot_b as float)) AS legal_hol_ot_b, SUM(cast(ep.special_hol_ot_b as float)) AS special_hol_ot_b, SUM(cast(ep.legal_hol_pay_b as float)) AS legal_hol_pay_b, SUM(cast(ep.spl_hol_pay_b as float)) AS spl_hol_pay_b, SUM(ep.sss_cont_a) AS sss_cont_a, SUM(ep.philhealth_cont_a) AS philhealth_cont_a, SUM(ep.pag_ibig_a) AS pag_ibig_a, SUM(ep.other_deduction) AS other_deduction, SUM(ep.others) AS others, SUM(ep.advances_loans) AS advances_loans, SUM(ep.other_earnings) AS other_earnings,SUM(ep.w_tax) AS w_tax, ep.empid, e.pay_rate, e.department, d.dept_name, e.firstname, e.lastname, e.mi FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee e ON e.empid=ep.empid LEFT JOIN rssys.hr_payrollpariod pp ON pp.pay_code=ep.ppid LEFT JOIN rssys.hr_department d ON e.department=d.deptid WHERE 1=1 " + WHERE + " GROUP BY ep.empid, e.pay_rate, e.department, d.dept_name, e.firstname, e.lastname, e.mi ORDER BY ep.empid";

            data.strSQL = strSQL;
            data.dept_name = dept_name;
            data.fy = fy;
            data.mo = mofrm + " - " + moto;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_payslip(String payperiod_code, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to, String empid, String byRateType, String sortBy, Boolean isEmpOnly, Boolean isDetailed)
        {
            String WHERE = "", ORDERBY = "", strSQL = "";
            String dept_name = "";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(payperiod_code))
            {
                WHERE += " AND ep.ppid='" + payperiod_code + "'";
            }
            if (!String.IsNullOrEmpty(empid))
            {
                WHERE += " AND ep.empid='" + empid + "'";
            }

            if (byRateType == "Monthly paid")
            {
                WHERE += " AND e.rate_type='M'";
            }
            else if (byRateType == "Daily paid")
            {
                WHERE += " AND e.rate_type='D'";
            }

            if (sortBy == "By department/section")
            {
                ORDERBY += "ORDER BY e.department, e.section";
            }
            else if (sortBy == "By payment type")
            {
                ORDERBY += "ORDER BY e.empid";
            }
            else if (sortBy == "By department/employee")
            {
                ORDERBY += "ORDER BY e.department, e.empid";
            }

            strSQL = "SELECT ep.*, d.dept_name, to_char(pp.date_from,'YYYY-MM-DD') AS date_from, to_char(pp.date_to,'YYYY-MM-DD') AS date_to, to_char(pp.date_from,'Month DD, YYYY') AS date_from_desc, to_char(pp.date_to,'Month DD, YYYY') AS date_to_desc, e.firstname, e.lastname, e.mi, e.tax_bracket FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee e ON e.empid=ep.empid LEFT JOIN rssys.hr_payrollpariod pp ON pp.pay_code=ep.ppid LEFT JOIN rssys.hr_department d ON e.department=d.deptid WHERE 1=1 " + WHERE + " " + ORDERBY + "";

            data.strSQL = strSQL;
            //rtxt_msg.Text = data.strSQL;
            //MessageBox.Show(data.strSQL);
            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_payroll_summary(String payperiod_code, String depttxt_frm, String deptcode_frm, String depttxt_to, String deptcode_to, String empid, String byRateType, String sortBy, Boolean isEmpOnly, Boolean isDetailed)
        {
            String WHERE = "", ORDERBY = "", strSQL = "";
            String dept_name = "";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(payperiod_code))
            {
                WHERE += " AND ep.ppid='" + payperiod_code + "'";
            }
            if (!String.IsNullOrEmpty(empid))
            {
                WHERE += " AND ep.empid='" + empid + "'";
            }

            if (byRateType == "Monthly paid")
            {
                WHERE += " AND e.rate_type='M'";
            }
            else if (byRateType == "Daily paid")
            {
                WHERE += " AND e.rate_type='D'";
            }

            if (sortBy == "By department/section")
            {
                ORDERBY += "ORDER BY e.department, e.section";
            }
            else if (sortBy == "By payment type")
            {
                ORDERBY += "ORDER BY e.empid";
            }
            else if (sortBy == "By department/employee")
            {
                ORDERBY += "ORDER BY e.department, e.empid";
            }

            strSQL = "SELECT ep.*, e.pay_rate, e.department, d.dept_name, to_char(pp.date_from,'YYYY-MM-DD') AS date_from, to_char(pp.date_to,'YYYY-MM-DD') AS date_to, to_char(pp.date_from,'Month DD, YYYY') AS date_from_desc, to_char(pp.date_to,'Month DD, YYYY') AS date_to_desc, e.firstname, e.lastname, e.mi, e.tax_bracket FROM rssys.hr_emp_payroll ep LEFT JOIN rssys.hr_employee e ON e.empid=ep.empid LEFT JOIN rssys.hr_payrollpariod pp ON pp.pay_code=ep.ppid LEFT JOIN rssys.hr_department d ON e.department=d.deptid WHERE 1=1 " + WHERE + " " + ORDERBY + "";

            data.strSQL = strSQL;
            data.report_type = byRateType;

            pnl_rpt_option_header.Hide();
            print();
        }

        public void print_employee_dtr(String deptcode_frm, String depttxt_frm, String deptcode_to, String depttxt_to, String empid, String empname, String dtp_from, String dtp_to)
        {
            String WHERE = "", WHERE2 = "", ORDERBY = "", strSQL = "";
            String dept_name = "";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }

            if (!String.IsNullOrEmpty(empid))
            {
                WHERE2 = " AND empid='" + empid + "'";
                WHERE += " AND e.empid='" + empid + "'";
            }

            strSQL = "SELECT to_char(pp.gen_date,'YYYY-MM-DD') AS gen_date, to_char(pp.gen_date,'DAY') AS mon, CASE WHEN dy.day=dayoff1 OR dy.day=dayoff2 THEN 'DO' ELSE '' END AS dayoff, e.empid, e.lastname||','||e.firstname||' '||e.mi AS empname, e.lastname, e.firstname, e.mi, e.dept_name, e.department, to_char(CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN total_hr_sat ELSE  total_hr END,'hh24:mi') AS total_hr, to_char(wk.work_hr,'hh24:mi') AS work_hr, wk.work_date, wk.timein, wk.timeout, CASE WHEN COALESCE(wk.empid,'')<>'' OR (dy.day=dayoff1 OR dy.day=dayoff2) THEN '0' ELSE '1' END AS absent,  CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN      CASE WHEN wk.timein::time-e.sched_timein_sat::time>'00:00'::time THEN to_char(wk.timein::time-e.sched_timein_sat::time,'hh24:mi') ELSE '' END        ELSE CASE WHEN wk.timein::time-e.sched_timein::time>'00:00'::time THEN to_char(wk.timein::time-e.sched_timein::time,'hh24:mi') ELSE '' END END AS late, CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN      CASE WHEN wk.timeout::time-e.sched_timeout_sat::time>'00:00'::time THEN to_char(wk.timeout::time-e.sched_timeout_sat::time,'hh24:mi') ELSE '' END       ELSE CASE WHEN wk.timeout::time-e.sched_timeout::time>'00:00'::time THEN to_char(wk.timeout::time-e.sched_timeout::time,'hh24:mi') ELSE '' END END AS undertime, CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN       CASE WHEN e.sched_timeout_sat::time-wk.timeout::time>'00:00'::time THEN to_char(e.sched_timeout_sat::time-wk.timeout::time,'hh24:mi') ELSE '' END       ELSE CASE WHEN e.sched_timeout::time-wk.timeout::time>'00:00'::time THEN to_char(e.sched_timeout::time-wk.timeout::time,'hh24:mi') ELSE '' END END AS overtime, lv.description AS leave_desc, lv.leave_type FROM (SELECT DISTINCT gen.gen_date::date FROM    (SELECT generate_series('" + dtp_from + "'::date, '" + dtp_to + "'::date, '1 day'::interval) AS gen_date)       gen JOIN rssys.hr_payrollpariod pp ON (gen.gen_date BETWEEN pp.date_from AND pp.date_to) ORDER BY gen_date) pp JOIN (SELECT e.*, d.dept_name, ss.am_timein, ss.am_timeout, ss.pm_timein, ss.pm_timeout, ssat.am_timein AS am_timein_sat, ssat.am_timeout AS am_timeout_sat, ssat.pm_timein AS pm_timein_sat, ssat.pm_timeout AS pm_timeout_sat,  (CASE WHEN COALESCE(ss.pm_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ss.pm_timeout,'00:00') END::time-CASE WHEN COALESCE(ss.pm_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ss.pm_timein,'00:00') END::time) + (CASE WHEN COALESCE(ss.am_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ss.am_timeout,'00:00') END::time-CASE WHEN COALESCE(ss.am_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ss.am_timein,'00:00') END::time) AS  total_hr, (CASE WHEN COALESCE(ssat.pm_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.pm_timeout,'00:00') END::time-CASE WHEN COALESCE(ssat.pm_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.pm_timein,'00:00') END::time) + (CASE WHEN COALESCE(ssat.am_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.am_timeout,'00:00') END::time-CASE WHEN COALESCE(ssat.am_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.am_timein,'00:00') END::time) AS  total_hr_sat, CASE WHEN COALESCE(ss.am_timeout,'')<>'' THEN ss.am_timein ELSE ss.pm_timein END sched_timein, CASE WHEN COALESCE(ss.pm_timeout,'')<>'' THEN ss.pm_timeout ELSE ss.am_timeout END sched_timeout, CASE WHEN COALESCE(ssat.am_timeout,'')<>'' THEN ssat.am_timein ELSE ssat.pm_timein END sched_timein_sat, CASE WHEN COALESCE(ssat.pm_timeout,'')<>'' THEN ssat.pm_timeout ELSE ssat.am_timeout END sched_timeout_sat FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON e.department=d.deptid LEFT JOIN rssys.hr_shift_schedule ss ON ss.code= e.shift_sched LEFT JOIN rssys.hr_shift_schedule ssat ON ssat.code=e.shift_sched  WHERE 1=1  " + WHERE + ") e ON (e.date_hired<>'1900-01-01' AND e.date_hired<=pp.gen_date)  LEFT JOIN (SELECT t.*, CASE WHEN maxam<>'00:00' AND maxap<>'00:00' AND minap<>'00:00' AND minam<>'00:00' THEN  maxam::time-minam::time +  to_char(maxap::time-minap::time,'hh24:mi')::interval WHEN maxam='00:00' AND maxap='00:00' AND minap<>'00:00' AND minam<>'00:00' THEN minap::time-minam::time WHEN minam='00:00' AND maxam='00:00' AND maxap<>'00:00' THEN  maxap::time-minap::time WHEN minap='00:00' AND maxap='00:00' AND maxam<>'00:00' THEN  maxam::time-minam::time ELSE '00:00'::interval END AS work_hr, minam AS timein, CASE WHEN maxap='00:00' THEN minap ELSE maxap END timeout FROM (SELECT t.empid, t.work_date, COALESCE(tam.min,'00:00') AS minam, CASE WHEN tam.min=tam.max THEN '00:00' ELSE COALESCE(tam.max,'00:00') END AS maxam, COALESCE(tap.min,'00:00') AS minap, CASE WHEN tap.min=tap.max THEN '00:00' ELSE COALESCE(tap.max,'00:00') END AS maxap FROM (SELECT DISTINCT work_date, empid FROM rssys.hr_tito2) t LEFT JOIN (SELECT DISTINCT work_date, MIN(time_log), MAX(time_log) FROM rssys.hr_tito2 WHERE TRIM(to_char(time_log::time,'AM'))='AM' GROUP BY work_date) tam ON tam.work_date=t.work_date LEFT JOIN (SELECT DISTINCT work_date, MIN(time_log), MAX(time_log) FROM rssys.hr_tito2 WHERE TRIM(to_char(time_log::time,'AM'))='PM' GROUP BY work_date) tap ON tap.work_date=t.work_date) t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid ) wk ON (wk.work_date=pp.gen_date AND e.empid=wk.empid) LEFT JOIN rssys.hr_days dy ON (dy.dayname=TRIM(to_char(pp.gen_date,'DAY'))) LEFT JOIN (SELECT lv.*, lt.description FROM rssys.hr_leaves lv LEFT JOIN rssys.hr_leave_type lt ON (lt.code=lv.leave_type) WHERE 1=1 " + WHERE2 + ") lv ON (e.empid=lv.empid AND pp.gen_date BETWEEN lv.leave_from AND lv.leave_to)           ORDER BY e.empid, pp.gen_date";

            data.strSQL = strSQL;
            data.dept_name = dept_name;
            data.date = dtp_from + " to " + dtp_to;

            pnl_rpt_option_header.Hide();
            print();
        }

        public void print_absencelateundertime(String deptcode_frm, String depttxt_frm, String deptcode_to, String depttxt_to, String empid, String empname, String dtp_from, String dtp_to)
        {
            String WHERE = "", WHERE2 = "", ORDERBY = "", strSQL = "";
            String dept_name = "";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }

            if (!String.IsNullOrEmpty(empid))
            {
                WHERE2 = " AND empid='" + empid + "'";
                WHERE += " AND e.empid='" + empid + "'";
            }

            strSQL = "SELECT to_char(pp.gen_date,'YYYY-MM-DD') AS gen_date, to_char(pp.gen_date,'DAY') AS mon, CASE WHEN dy.day=dayoff1 OR dy.day=dayoff2 THEN 'DO' ELSE '' END AS dayoff, e.empid, e.lastname||','||e.firstname||' '||e.mi AS empname, e.lastname, e.firstname, e.mi, e.dept_name, e.department, to_char(CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN total_hr_sat ELSE  total_hr END,'hh24:mi') AS total_hr, to_char(wk.work_hr,'hh24:mi') AS work_hr, wk.work_date, wk.timein, wk.timeout, CASE WHEN COALESCE(wk.empid,'')<>'' OR (dy.day=dayoff1 OR dy.day=dayoff2) THEN '0' ELSE '1' END AS absent,  CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN      CASE WHEN wk.timein::time-e.sched_timein_sat::time>'00:00'::time THEN to_char(wk.timein::time-e.sched_timein_sat::time,'hh24:mi') ELSE '' END        ELSE CASE WHEN wk.timein::time-e.sched_timein::time>'00:00'::time THEN to_char(wk.timein::time-e.sched_timein::time,'hh24:mi') ELSE '' END END AS late, CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN      CASE WHEN wk.timeout::time-e.sched_timeout_sat::time>'00:00'::time THEN to_char(wk.timeout::time-e.sched_timeout_sat::time,'hh24:mi') ELSE '' END       ELSE CASE WHEN wk.timeout::time-e.sched_timeout::time>'00:00'::time THEN to_char(wk.timeout::time-e.sched_timeout::time,'hh24:mi') ELSE '' END END AS undertime, CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN       CASE WHEN e.sched_timeout_sat::time-wk.timeout::time>'00:00'::time THEN to_char(e.sched_timeout_sat::time-wk.timeout::time,'hh24:mi') ELSE '' END       ELSE CASE WHEN e.sched_timeout::time-wk.timeout::time>'00:00'::time THEN to_char(e.sched_timeout::time-wk.timeout::time,'hh24:mi') ELSE '' END END AS overtime, lv.description AS leave_desc, lv.leave_type FROM (SELECT DISTINCT gen.gen_date::date FROM    (SELECT generate_series('" + dtp_from + "'::date, '" + dtp_to + "'::date, '1 day'::interval) AS gen_date)       gen JOIN rssys.hr_payrollpariod pp ON (gen.gen_date BETWEEN pp.date_from AND pp.date_to) ORDER BY gen_date) pp JOIN (SELECT e.*, d.dept_name, ss.am_timein, ss.am_timeout, ss.pm_timein, ss.pm_timeout, ssat.am_timein AS am_timein_sat, ssat.am_timeout AS am_timeout_sat, ssat.pm_timein AS pm_timein_sat, ssat.pm_timeout AS pm_timeout_sat,  (CASE WHEN COALESCE(ss.pm_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ss.pm_timeout,'00:00') END::time-CASE WHEN COALESCE(ss.pm_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ss.pm_timein,'00:00') END::time) + (CASE WHEN COALESCE(ss.am_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ss.am_timeout,'00:00') END::time-CASE WHEN COALESCE(ss.am_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ss.am_timein,'00:00') END::time) AS  total_hr, (CASE WHEN COALESCE(ssat.pm_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.pm_timeout,'00:00') END::time-CASE WHEN COALESCE(ssat.pm_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.pm_timein,'00:00') END::time) + (CASE WHEN COALESCE(ssat.am_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.am_timeout,'00:00') END::time-CASE WHEN COALESCE(ssat.am_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.am_timein,'00:00') END::time) AS  total_hr_sat, CASE WHEN COALESCE(ss.am_timeout,'')<>'' THEN ss.am_timein ELSE ss.pm_timein END sched_timein, CASE WHEN COALESCE(ss.pm_timeout,'')<>'' THEN ss.pm_timeout ELSE ss.am_timeout END sched_timeout, CASE WHEN COALESCE(ssat.am_timeout,'')<>'' THEN ssat.am_timein ELSE ssat.pm_timein END sched_timein_sat, CASE WHEN COALESCE(ssat.pm_timeout,'')<>'' THEN ssat.pm_timeout ELSE ssat.am_timeout END sched_timeout_sat FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON e.department=d.deptid LEFT JOIN rssys.hr_shift_schedule ss ON ss.code= e.shift_sched LEFT JOIN rssys.hr_shift_schedule ssat ON ssat.code=e.shift_sched  WHERE 1=1  " + WHERE + ") e ON (e.date_hired<>'1900-01-01' AND e.date_hired<=pp.gen_date)  LEFT JOIN (SELECT t.*, CASE WHEN maxam<>'00:00' AND maxap<>'00:00' AND minap<>'00:00' AND minam<>'00:00' THEN  maxam::time-minam::time +  to_char(maxap::time-minap::time,'hh24:mi')::interval WHEN maxam='00:00' AND maxap='00:00' AND minap<>'00:00' AND minam<>'00:00' THEN minap::time-minam::time WHEN minam='00:00' AND maxam='00:00' AND maxap<>'00:00' THEN  maxap::time-minap::time WHEN minap='00:00' AND maxap='00:00' AND maxam<>'00:00' THEN  maxam::time-minam::time ELSE '00:00'::interval END AS work_hr, minam AS timein, CASE WHEN maxap='00:00' THEN minap ELSE maxap END timeout FROM (SELECT t.empid, t.work_date, COALESCE(tam.min,'00:00') AS minam, CASE WHEN tam.min=tam.max THEN '00:00' ELSE COALESCE(tam.max,'00:00') END AS maxam, COALESCE(tap.min,'00:00') AS minap, CASE WHEN tap.min=tap.max THEN '00:00' ELSE COALESCE(tap.max,'00:00') END AS maxap FROM (SELECT DISTINCT work_date, empid FROM rssys.hr_tito2) t LEFT JOIN (SELECT DISTINCT work_date, MIN(time_log), MAX(time_log) FROM rssys.hr_tito2 WHERE TRIM(to_char(time_log::time,'AM'))='AM' GROUP BY work_date) tam ON tam.work_date=t.work_date LEFT JOIN (SELECT DISTINCT work_date, MIN(time_log), MAX(time_log) FROM rssys.hr_tito2 WHERE TRIM(to_char(time_log::time,'AM'))='PM' GROUP BY work_date) tap ON tap.work_date=t.work_date) t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid ) wk ON (wk.work_date=pp.gen_date AND e.empid=wk.empid) LEFT JOIN rssys.hr_days dy ON (dy.dayname=TRIM(to_char(pp.gen_date,'DAY'))) LEFT JOIN (SELECT lv.*, lt.description FROM rssys.hr_leaves lv LEFT JOIN rssys.hr_leave_type lt ON (lt.code=lv.leave_type) WHERE 1=1 " + WHERE2 + ") lv ON (e.empid=lv.empid AND pp.gen_date BETWEEN lv.leave_from AND lv.leave_to)    WHERE (CASE WHEN dy.day=dayoff1 OR dy.day=dayoff2 THEN 'DO' ELSE '' END<>'DO')       ORDER BY e.empid, pp.gen_date";

            data.strSQL = strSQL;
            data.dept_name = dept_name;
            data.date = dtp_from + " to " + dtp_to;

            pnl_rpt_option_header.Hide();
            print();
        }
        public void print_daily_timelog(String deptcode_frm, String depttxt_frm, String deptcode_to, String depttxt_to, String empid, String empname, String dtp_from, String dtp_to)
        {
            String WHERE = "", WHERE2 = "", ORDERBY = "", strSQL = "";
            String dept_name = "";

            if (!String.IsNullOrEmpty(deptcode_frm))
            {
                WHERE = " AND '" + deptcode_frm + "' <= d.deptid ";
                dept_name = depttxt_frm + " To Last";
            }
            if (!String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND '" + deptcode_to + "' >= d.deptid ";
                dept_name = "First To " + depttxt_to;
            }
            if (!String.IsNullOrEmpty(deptcode_frm) && !String.IsNullOrEmpty(deptcode_to))
            {
                WHERE = " AND d.deptid BETWEEN '" + deptcode_frm + "'  AND '" + deptcode_to + "' ";
                if (deptcode_frm == deptcode_to) dept_name = depttxt_frm;
                else dept_name = depttxt_frm + " To " + depttxt_to;
            }

            if (!String.IsNullOrEmpty(empid))
            {
                WHERE2 = " AND empid='" + empid + "'";
                WHERE += " AND e.empid='" + empid + "'";
            }
            //(SELECT generate_series('1900-01-01'::date, '1900-02-01'::date, '1 day'::interval) AS gen_date)
            //(SELECT generate_series(MIN(date_from), MAX(date_to), '1 day'::interval) AS gen_date FROM rssys.hr_payrollpariod)
            //strSQL = "SELECT to_char(pp.gen_date,'YYYY-MM-DD') AS gen_date, to_char(pp.gen_date,'DAY') AS mon, CASE WHEN dy.day=dayoff1 OR dy.day=dayoff2 THEN 'DO' ELSE '' END AS dayoff, e.empid, e.lastname||','||e.firstname||' '||e.mi AS empname, e.lastname, e.firstname, e.mi, e.dept_name, e.department, to_char(shift_sched_to::time - shift_sched_from::time,'hh24:mi') AS work_hr,wk.work_date,wk.timein,wk.timeout, CASE WHEN COALESCE(wk.empid,'')<>'' OR (dy.day=dayoff1 OR dy.day=dayoff2) THEN '0' ELSE '1' END AS absent,  CASE WHEN wk.timein::time-e.shift_sched_from::time>'00:00'::time THEN to_char(wk.timein::time-e.shift_sched_from::time,'hh24:mi') ELSE '' END AS late, CASE WHEN wk.timeout::time-e.shift_sched_to::time>'00:00'::time THEN to_char(wk.timeout::time-e.shift_sched_to::time,'hh24:mi') ELSE '' END AS undertime, CASE WHEN e.shift_sched_to::time-wk.timeout::time>'00:00'::time THEN to_char(e.shift_sched_to::time-wk.timeout::time,'hh24:mi') ELSE '' END AS overtime ,lv.description AS leave_desc, lv.leave_type               FROM (SELECT DISTINCT gen.gen_date::date FROM    (SELECT generate_series('" + dtp_from + "'::date, '" + dtp_to + "'::date, '1 day'::interval) AS gen_date)       gen JOIN rssys.hr_payrollpariod pp ON (gen.gen_date BETWEEN pp.date_from AND pp.date_to) ORDER BY gen_date) pp JOIN (SELECT e.*, d.dept_name FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON e.department=d.deptid WHERE 1=1 " + WHERE + ") e ON (e.date_hired<>'1900-01-01' AND e.date_hired<=pp.gen_date)  LEFT JOIN (SELECT DISTINCT e.empid,work_date,(SELECT MIN(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='I' AND empid=t.empid) AS timein, (SELECT MAX(time_log) FROM rssys.hr_tito2 st WHERE work_date=t.work_date AND status='O' AND empid=t.empid) AS timeout FROM rssys.hr_tito2 t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid ) wk ON (wk.work_date=pp.gen_date AND e.empid=wk.empid) LEFT JOIN rssys.hr_days dy ON (dy.dayname=TRIM(to_char(pp.gen_date,'DAY'))) LEFT JOIN (SELECT lv.*, lt.description FROM rssys.hr_leaves lv LEFT JOIN rssys.hr_leave_type lt ON (lt.code=lv.leave_type) WHERE 1=1 " + WHERE2 + ") lv ON (e.empid=lv.empid AND pp.gen_date BETWEEN lv.leave_from AND lv.leave_to)           ORDER BY e.empid, pp.gen_date";


            strSQL = "SELECT to_char(pp.gen_date,'YYYY-MM-DD') AS gen_date, to_char(pp.gen_date,'DAY') AS mon, CASE WHEN dy.day=dayoff1 OR dy.day=dayoff2 THEN 'DO' ELSE '' END AS dayoff, e.empid, e.lastname||','||e.firstname||' '||e.mi AS empname, e.lastname, e.firstname, e.mi, e.dept_name, e.department, to_char(CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN total_hr_sat ELSE  total_hr END,'hh24:mi') AS total_hr, to_char(wk.work_hr,'hh24:mi') AS work_hr, wk.work_date, wk.timein, wk.timeout, CASE WHEN COALESCE(wk.empid,'')<>'' OR (dy.day=dayoff1 OR dy.day=dayoff2) THEN '0' ELSE '1' END AS absent,  CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN      CASE WHEN wk.timein::time-e.sched_timein_sat::time>'00:00'::time THEN to_char(wk.timein::time-e.sched_timein_sat::time,'hh24:mi') ELSE '' END        ELSE CASE WHEN wk.timein::time-e.sched_timein::time>'00:00'::time THEN to_char(wk.timein::time-e.sched_timein::time,'hh24:mi') ELSE '' END END AS late, CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN      CASE WHEN wk.timeout::time-e.sched_timeout_sat::time>'00:00'::time THEN to_char(wk.timeout::time-e.sched_timeout_sat::time,'hh24:mi') ELSE '' END       ELSE CASE WHEN wk.timeout::time-e.sched_timeout::time>'00:00'::time THEN to_char(wk.timeout::time-e.sched_timeout::time,'hh24:mi') ELSE '' END END AS undertime, CASE WHEN trim(to_char(pp.gen_date,'day'))='saturday' THEN       CASE WHEN e.sched_timeout_sat::time-wk.timeout::time>'00:00'::time THEN to_char(e.sched_timeout_sat::time-wk.timeout::time,'hh24:mi') ELSE '' END       ELSE CASE WHEN e.sched_timeout::time-wk.timeout::time>'00:00'::time THEN to_char(e.sched_timeout::time-wk.timeout::time,'hh24:mi') ELSE '' END END AS overtime, lv.description AS leave_desc, lv.leave_type FROM (SELECT DISTINCT gen.gen_date::date FROM    (SELECT generate_series('" + dtp_from + "'::date, '" + dtp_to + "'::date, '1 day'::interval) AS gen_date)       gen JOIN rssys.hr_payrollpariod pp ON (gen.gen_date BETWEEN pp.date_from AND pp.date_to) ORDER BY gen_date) pp JOIN (SELECT e.*, d.dept_name, ss.am_timein, ss.am_timeout, ss.pm_timein, ss.pm_timeout, ssat.am_timein AS am_timein_sat, ssat.am_timeout AS am_timeout_sat, ssat.pm_timein AS pm_timein_sat, ssat.pm_timeout AS pm_timeout_sat,  (CASE WHEN COALESCE(ss.pm_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ss.pm_timeout,'00:00') END::time-CASE WHEN COALESCE(ss.pm_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ss.pm_timein,'00:00') END::time) + (CASE WHEN COALESCE(ss.am_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ss.am_timeout,'00:00') END::time-CASE WHEN COALESCE(ss.am_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ss.am_timein,'00:00') END::time) AS  total_hr, (CASE WHEN COALESCE(ssat.pm_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.pm_timeout,'00:00') END::time-CASE WHEN COALESCE(ssat.pm_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.pm_timein,'00:00') END::time) + (CASE WHEN COALESCE(ssat.am_timeout,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.am_timeout,'00:00') END::time-CASE WHEN COALESCE(ssat.am_timein,'00:00')='' THEN '00:00' ELSE COALESCE(ssat.am_timein,'00:00') END::time) AS  total_hr_sat, CASE WHEN COALESCE(ss.am_timeout,'')<>'' THEN ss.am_timein ELSE ss.pm_timein END sched_timein, CASE WHEN COALESCE(ss.pm_timeout,'')<>'' THEN ss.pm_timeout ELSE ss.am_timeout END sched_timeout, CASE WHEN COALESCE(ssat.am_timeout,'')<>'' THEN ssat.am_timein ELSE ssat.pm_timein END sched_timein_sat, CASE WHEN COALESCE(ssat.pm_timeout,'')<>'' THEN ssat.pm_timeout ELSE ssat.am_timeout END sched_timeout_sat FROM rssys.hr_employee e LEFT JOIN rssys.hr_department d ON e.department=d.deptid LEFT JOIN rssys.hr_shift_schedule ss ON ss.code= e.shift_sched LEFT JOIN rssys.hr_shift_schedule ssat ON ssat.code=e.shift_sched  WHERE 1=1  " + WHERE + ") e ON (e.date_hired<>'1900-01-01' AND e.date_hired<=pp.gen_date)  LEFT JOIN (SELECT t.*, CASE WHEN maxam<>'00:00' AND maxap<>'00:00' AND minap<>'00:00' AND minam<>'00:00' THEN  maxam::time-minam::time +  to_char(maxap::time-minap::time,'hh24:mi')::interval WHEN maxam='00:00' AND maxap='00:00' AND minap<>'00:00' AND minam<>'00:00' THEN minap::time-minam::time WHEN minam='00:00' AND maxam='00:00' AND maxap<>'00:00' THEN  maxap::time-minap::time WHEN minap='00:00' AND maxap='00:00' AND maxam<>'00:00' THEN  maxam::time-minam::time ELSE '00:00'::interval END AS work_hr, minam AS timein, CASE WHEN maxap='00:00' THEN minap ELSE maxap END timeout FROM (SELECT t.empid, t.work_date, COALESCE(tam.min,'00:00') AS minam, CASE WHEN tam.min=tam.max THEN '00:00' ELSE COALESCE(tam.max,'00:00') END AS maxam, COALESCE(tap.min,'00:00') AS minap, CASE WHEN tap.min=tap.max THEN '00:00' ELSE COALESCE(tap.max,'00:00') END AS maxap FROM (SELECT DISTINCT work_date, empid FROM rssys.hr_tito2) t LEFT JOIN (SELECT DISTINCT work_date, MIN(time_log), MAX(time_log) FROM rssys.hr_tito2 WHERE TRIM(to_char(time_log::time,'AM'))='AM' GROUP BY work_date) tam ON tam.work_date=t.work_date LEFT JOIN (SELECT DISTINCT work_date, MIN(time_log), MAX(time_log) FROM rssys.hr_tito2 WHERE TRIM(to_char(time_log::time,'AM'))='PM' GROUP BY work_date) tap ON tap.work_date=t.work_date) t LEFT JOIN rssys.hr_employee e ON t.empid=e.empid ) wk ON (wk.work_date=pp.gen_date AND e.empid=wk.empid) LEFT JOIN rssys.hr_days dy ON (dy.dayname=TRIM(to_char(pp.gen_date,'DAY'))) LEFT JOIN (SELECT lv.*, lt.description FROM rssys.hr_leaves lv LEFT JOIN rssys.hr_leave_type lt ON (lt.code=lv.leave_type) WHERE 1=1 " + WHERE2 + ") lv ON (e.empid=lv.empid AND pp.gen_date BETWEEN lv.leave_from AND lv.leave_to)           ORDER BY e.empid, pp.gen_date";

            data.strSQL = strSQL;
            data.dept_name = dept_name;
            data.date = dtp_from + " to " + dtp_to;

            pnl_rpt_option_header.Hide();
            print();
        }
    }
}
