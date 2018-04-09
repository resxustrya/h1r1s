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
    public partial class Main : Form
    {
        thisDatabase db;
        GlobalClass gc;
        GlobalMethod gm;
        public Color colormain;
        public Color color2;

        public Main()
        {
            InitializeComponent();
            color2 = new Color();
            colormain = new Color();
            db = new thisDatabase();
            gm = new GlobalMethod();
            GlobalClass.user_fullname = db.getFullName();
            gc = new GlobalClass();
            String version = "";

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = string.Format("Version {0}", ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4));
            }

            this.Text = this.Text + " " + version + " - Server IP " + GlobalClass.server_ip ;

            lbl_m99company.Text = db.get_m99comp_name();
            lbl_trnxdate.Text = db.get_systemdate("ddd, MM/dd/yy");
            lbl_user.Text = GlobalClass.username;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            z_Notification jsi = new z_Notification();

            jsi.MdiParent = this;
            lbl_modname.Text = btn_7.Text;
            btn_7.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;
            jsi.Show();
        }

        private void closechild()
        {
            foreach (Form f in this.MdiChildren)
            {
                f.Close();
            }
        }

        private void btn_color_reset()
        {
            btn_1.BackColor = menuStrip1.BackColor;
            btn_2.BackColor = menuStrip1.BackColor;
            btn_3.BackColor = menuStrip1.BackColor;
            btn_4.BackColor = menuStrip1.BackColor;
            btn_5.BackColor = menuStrip1.BackColor;
            btn_6.BackColor = menuStrip1.BackColor;
            btn_7.BackColor = menuStrip1.BackColor;
            btn_8.BackColor = menuStrip1.BackColor;
            btn_9.BackColor = menuStrip1.BackColor;
            btn_10.BackColor = menuStrip1.BackColor;
        }

        public String get_modname()
        {
            return lbl_modname.Text;
        }

        private void open_logbox()
        {
            t_LogBox jsi = new t_LogBox();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_1.Text;
            btn_1.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_uploadtimelogfile()
        {
            t_UploadLogsFile jsi = new t_UploadLogsFile();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_1.Text;
            btn_1.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_batchtimelog()
        {
            t_BatchTimeLog jsi = new t_BatchTimeLog();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_2.Text;
            btn_2.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            
            jsi.Show();
        }

        private void open_empdtr()
        {
            t_emp_dtr jsi = new t_emp_dtr();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_3.Text;
            btn_3.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_leavesentry()
        {
            t_Leaves jsi = new t_Leaves();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = "Employee Leave Entry";//btn_4.Text;
            //btn_4.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_manpowerscheduleentry()
        {
            t_manpowerschedule jsi = new t_manpowerschedule();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = "Employee Schedule";//btn_5.Text;
            //btn_5.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_generatedtr()
        {
            t_generatedtr jsi = new t_generatedtr();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_4.Text;
            btn_4.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_emploan()
        {
            l_LoanEntry jsi = new l_LoanEntry();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_5.Text;
            btn_5.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }
        
        private void open_emploanledgercard()
        {
            l_LoanLedgerCard jsi = new l_LoanLedgerCard();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = "Employee Loan History";
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_generatepayroll()
        {
            p_GeneratePayroll jsi = new p_GeneratePayroll();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_9.Text;
            btn_9.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_viewgeneratedpayroll()
        {
            p_ViewGeneratedPayroll jsi = new p_ViewGeneratedPayroll();

            //closechild();
            //btn_color_reset();

            //jsi.MdiParent = this;

            //lbl_modname.Text = btn_10.Text;
            //btn_10.BackColor = panel2.BackColor;
            //colormain = panel2.BackColor;
            //color2 = Color.Peru;

            jsi.ShowDialog();
        }

        private void open_otherearnings()
        {
            p_OtherEarningEntry jsi = new p_OtherEarningEntry();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_6.Text;
            btn_6.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void open_otherdeduction()
        {
            p_OtherDeductionEntry jsi = new p_OtherDeductionEntry();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_8.Text;
            btn_8.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }




        private void btn_1_Click(object sender, EventArgs e)
        {
            //open_logbox();
            open_uploadtimelogfile();
        }

        private void btn_2_Click(object sender, EventArgs e)
        {
            open_batchtimelog();
        }

        private void btn_3_Click(object sender, EventArgs e)
        {
            open_empdtr();
        }

        private void btn_4_Click(object sender, EventArgs e)
        {
            open_generatedtr();
        }

        private void btn_5_Click(object sender, EventArgs e)
        {
            open_emploan();
        }

        private void btn_6_Click(object sender, EventArgs e)
        {
            open_otherearnings();
        }

        private void btn_8_Click(object sender, EventArgs e)
        {
            open_otherdeduction();
        }

        private void btn_9_Click(object sender, EventArgs e)
        {
            open_generatepayroll();
        }

        private void btn_10_Click(object sender, EventArgs e)
        {
            open_viewgeneratedpayroll();
        }

        private void departmentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Department frm = new m_Department();

            frm.ShowDialog();
        }

        private void departmenSectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Section frm = new m_Section();

            frm.ShowDialog();
        }

        private void positionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Position frm = new m_Position();

            frm.ShowDialog();
        }

        private void fmi_m600_Click(object sender, EventArgs e)
        {

        }

        private void payrollPeriodToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_payrollperiod frm = new m_payrollperiod();

            frm.ShowDialog();
        }

        private void witholdingTaxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_WTax frm = new m_WTax();

            frm.ShowDialog();
        }

        private void pHILHEALTHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_PhilHealth frm = new m_PhilHealth();

            frm.ShowDialog();
        }

        private void hDMFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_HDMF frm = new m_HDMF();

            frm.ShowDialog();
        }

        private void loanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Loan frm = new m_Loan();

            frm.ShowDialog();
        }

        private void otherEarningsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_OtherEarnings frm = new m_OtherEarnings();

            frm.ShowDialog();
        }

        private void otherDeductionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_OtherDeductions frm = new m_OtherDeductions();

            frm.ShowDialog();
        }

        private void holidaysToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Holidays frm = new m_Holidays();

            frm.ShowDialog();
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void btn_7_Click(object sender, EventArgs e)
        {
            z_Notification jsi = new z_Notification();

            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text = btn_7.Text;
            btn_7.BackColor = panel2.BackColor;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }        

        private void logBoxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_logbox();
        }

        private void textBlastToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_uploadtimelogfile();
        }

        private void batchTimeLogEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_batchtimelog();
        }

        private void leavesEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_leavesentry();
        }

        private void manpowerScheduleEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_manpowerscheduleentry();
        }

        private void generateDTRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_generatedtr();
        }

        private void roomAvailabilityToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_emploan();
        }

        private void generatePayrollToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_generatepayroll();
        }

        private void viewGeneratedPayrollToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_viewgeneratedpayroll();
        }

        private void employeeLedgerCardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void fmi_m004_Click(object sender, EventArgs e)
        {

        }

        private void fmi_m004_Click_1(object sender, EventArgs e)
        {
            m_employee frm = new m_employee();
            frm.ShowDialog();
        }

        private void fmi_m000_Click(object sender, EventArgs e)
        {
            
        }

        private void shiftScheduleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_ShiftSchedule ss = new m_ShiftSchedule();
            ss.ShowDialog();
        }

        private void employeeDTRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_empdtr();
        }

        private void payrollEntriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
           p_PayrollRegister jsi = new p_PayrollRegister();
           closechild();
           btn_color_reset();

           jsi.MdiParent = this;

           lbl_modname.Text = jsi.Text;
           colormain = panel2.BackColor;
           color2 = Color.Peru;

           jsi.Show();
        }

        private void otherEarningsEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_otherearnings();
        }

        private void otherDeductionsEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_otherdeduction();
        }

        private void payslipsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            p_Payslips jsi = new p_Payslips();
            closechild();
            btn_color_reset();

            jsi.MdiParent = this;

            lbl_modname.Text =jsi.Text;
            colormain = panel2.BackColor;
            color2 = Color.Peru;

            jsi.Show();
        }

        private void sSSTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_SSS frm = new m_SSS();
            frm.ShowDialog();
        }

        private void jobTitleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_jobtitle frm = new m_jobtitle();
            frm.ShowDialog();
        }

        private void businessUnitsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_business_units frm = new m_business_units();
            frm.ShowDialog();
        }

        private void contributionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_contribution_remittance frm = new m_contribution_remittance();
            frm.ShowDialog();
        }

        private void leaveTypesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Leaves frm = new m_Leaves();
            frm.ShowDialog();
        }

        private void loanEntryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            open_emploan();
        }

        private void employeeLoanHistoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open_emploanledgercard();
        }

        private void employeeRegisterToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void payrollSummaryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("P101");
            form.Show();
        }
        private void payslipToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("P102");
            form.Show();
        }

        private void employeeListingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("E103");
            form.Show();
        }

        private void otherIncomeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("M102");
            form.Show();
        }

        private void otherDeductionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("M103");
            form.Show();
        }

        private void otherDeductionPerEmployeeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("M104");
            form.Show();
        }

        private void dTRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("T101");
            form.Show();
        }
        private void dailyTimelogRecordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("T102");
            form.Show();
        }

        private void absencesLateAndTardinessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("T103");
            form.Show();
        }

        private void sSSContributionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("P202");
            form.Show();
        }

        private void philhealthContributionSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RPT_RES_entry2 form = new RPT_RES_entry2("P203");
            form.Show();
        }


    }
}
