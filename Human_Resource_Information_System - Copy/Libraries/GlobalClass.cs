using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Globalization;

namespace Human_Resource_Information_System
{
    class GlobalClass
    {
        private static string l_user = null;
        private static string l_userfullname = null;
        private static string l_branch = null;
        private static string l_schema = null;
        private static DataTable l_gdt = null;
        private static DataRow l_gdr = null;
        private static DataGridView l_gdgv = null;
        private static DataGridViewRow l_gdgvRow = null;
        private static string l_projcomp = null;
        private static Boolean l_DontSendToMain = false;
        private static String l_ipaddress = "";
        
        public static Boolean DontSendToMain
        {
            get { return l_DontSendToMain; }
            set { l_DontSendToMain = value; }
        }

        public static String server_ip
        {
            get { return l_ipaddress; }
            set { l_ipaddress = value; }
        }

        public static string projcompany
        {
            get { return l_projcomp; }
            set { l_projcomp = value; }
        }

        public static string username
        {
            get { return l_user; }
            set { l_user = value; }
        }

        public static string user_fullname
        {
            get { return l_userfullname; }
            set { l_userfullname = value; }
        }

        public static string branch
        {
            get { return l_branch; }
            set { l_branch = value; }
        }

        public static string schema
        {
            get { return l_schema; }
            set { l_schema = value; }
        }

        public static DataTable gdt
        {
            get { return l_gdt; }
            set { l_gdt = value; }
        }

        public static DataRow gdr
        {
            get { return l_gdr; }
            set { l_gdr = value; }
        }

        public static DataGridView gdgv
        {
            get { return l_gdgv; }
            set { l_gdgv = value; }
        }

        public static DataGridViewRow gdgvRow
        {
            get { return l_gdgvRow; }
            set { l_gdgvRow = value; }
        }

        public DataGridView CopyDataGridView(DataGridView dgv_org)
        {
            DataGridView dgv_copy = new DataGridView();
            try
            {
                if (dgv_copy.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn dgvc in dgv_org.Columns)
                    {
                        dgv_copy.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                    }
                }

                DataGridViewRow row = new DataGridViewRow();

                for (int i = 0; i < dgv_org.Rows.Count; i++)
                {
                    row = (DataGridViewRow)dgv_org.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dgv_org.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dgv_copy.Rows.Add(row);
                }
                dgv_copy.AllowUserToAddRows = false;
                dgv_copy.Refresh();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);//cf.ShowExceptionErrorMsg("Copy DataGridViw", ex);
            }

            return dgv_copy;
        }

        public void set_cbo_selectedvalue(ComboBox cbo, String selectedvalue)
        {
            //try
           // {
                if (String.IsNullOrEmpty(selectedvalue) == false)
                {
                    cbo.SelectedValue = selectedvalue;
                }
                else
                {
                    cbo.SelectedIndex = -1;
                }
            //}
            //catch (Exception) { cbo.SelectedIndex = -1; }
        }

        public void load_branch(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("branch", "code, name", "", " ORDER BY name ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "name";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_journal(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m05", "j_code, j_desc", "", " ORDER BY j_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "j_desc";
                cbo.ValueMember = "j_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_journal_code_asc(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m05", "j_code, j_desc", "", " ORDER BY j_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "j_desc";
                cbo.ValueMember = "j_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_openperiod(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("x03", "fy ||'-'|| mo AS mo, fy || ' - ' ||month_desc AS month_desc", "", " ORDER BY \"from\" ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "month_desc";
                cbo.ValueMember = "mo";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_userid(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("x08", "uid", "", " ORDER BY uid ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "uid";
                cbo.ValueMember = "uid";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_userfullname(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("x08", "uid, opr_name", "", " ORDER BY uid ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "opr_name";
                cbo.ValueMember = "uid";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_customer(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.get_customer_list();
                cbo.DataSource = dt;
                cbo.DisplayMember = "d_name";
                cbo.ValueMember = "d_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_m00(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m00", "code, name", "", " ORDER BY code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "name";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_m01(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m01", "mag_code, mag_desc", "", " ORDER BY mag_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "mag_desc";
                cbo.ValueMember = "mag_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_m02(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m02", "cmp_code, cmp_desc", "", " ORDER BY cmp_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "cmp_desc";
                cbo.ValueMember = "cmp_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_subaccount(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m03", "acc_code, acc_desc", "", " ORDER BY acc_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "acc_desc";
                cbo.ValueMember = "acc_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_accounttitle(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "", " ORDER BY at_desc ASC");

                db.CloseConn();
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_accounttitle_sl_only(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "sl='Y'", " ORDER BY at_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_accounttitle_payment_only(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "payment='Y'", " ORDER BY at_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_mop(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m10", "mp_code, mp_desc", "", " ORDER BY mp_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "mp_desc";
                cbo.ValueMember = "mp_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_collector(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("collector", "coll_code, coll_name", "", " ORDER BY coll_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "coll_name";
                cbo.ValueMember = "coll_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_accountingperiod(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("x04", "mo, month_desc", "", " ORDER BY mo ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "month_desc";
                cbo.ValueMember = "mo";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        //for customer ledger
        public void load_account_for_cust_ledger(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "dr_cr='D' AND sl='Y'", " ORDER BY at_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        //for supplier ledger
        public void load_account_for_sup_ledger(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "dr_cr='C' AND sl='Y'", " ORDER BY at_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_costcenter(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m08", "cc_code, cc_desc", "", " ORDER BY cc_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "cc_desc";
                cbo.ValueMember = "cc_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_subsidiaryname(ComboBox cbo, String accttitle_code)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();
                String WHERE = "";
                Boolean ispayable = false;

                ispayable = db.is_liabilities(accttitle_code);

                if (ispayable == true)
                {
                    if (String.IsNullOrEmpty(accttitle_code) == false)
                    {
                        WHERE = " at_code='" + accttitle_code + "'";
                    }

                    dt = db.QueryOnTableWithParams("m07", "c_code, c_name", WHERE, " ORDER BY c_name ASC");
                    cbo.DataSource = dt;
                    cbo.DisplayMember = "c_name";
                    cbo.ValueMember = "c_code";
                    cbo.SelectedIndex = -1;
                }
                else
                {
                    if (String.IsNullOrEmpty(accttitle_code) == false)
                    {
                        WHERE = " at_code='" + accttitle_code + "'";
                    }

                    dt = db.QueryOnTableWithParams("m06", "d_code, d_name", WHERE, " ORDER BY d_name ASC");
                    cbo.DataSource = dt;
                    cbo.DisplayMember = "d_name";
                    cbo.ValueMember = "d_code";
                    cbo.SelectedIndex = -1;
                }
            }
            catch (Exception) { }
        }

        public void load_subcostcenter(ComboBox cbo, String costcenter_code)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();
                String WHERE = "";

                if (String.IsNullOrEmpty(costcenter_code) == false)
                {
                    WHERE = " cc_code='" + costcenter_code + "'";
                }

                dt = db.QueryOnTableWithParams("subctr", "scc_code, scc_desc", WHERE, " ORDER BY scc_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "scc_desc";
                cbo.ValueMember = "scc_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_payee(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("payee", "payee", "", " ORDER BY payee ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "payee";
                cbo.ValueMember = "payee";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_terms(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m10", "mp_code, mp_desc", "", "");
                cbo.DataSource = dt;
                cbo.DisplayMember = "mp_desc";
                cbo.ValueMember = "mp_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_whouse(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("whouse", "whs_code, whs_desc", "", "");
                cbo.DataSource = dt;
                cbo.DisplayMember = "whs_desc";
                cbo.ValueMember = "whs_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_poinvoice_nonjrnlz(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("pinvhd", "inv_num, inv_num || ' - ' || supl_name || ' - ' || reference AS po_desc", "jrnlz!='Y' AND cancel='N'", "ORDER BY inv_num ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "po_desc";
                cbo.ValueMember = "inv_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_poinvoice_return_nonjrnlz(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("prethdr", "pret_num, pret_num || ' - ' || reference || ' - ' || supl_name  AS pret_desc", "(jrnlz='N' OR jrnlz is null) AND cancel is null", "ORDER BY pret_num ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "pret_desc";
                cbo.ValueMember = "pret_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_stktransfer_invoice(ComboBox cbo, String whs_code)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("rechdr", "rec_num, rec_num || ' - ' || trnx_date || ' - ' || reference  AS rec_desc", "whs_code='" + whs_code + "' AND trn_type='T' AND (cancel!='Y' OR cancel is null)", "ORDER BY rec_num ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "rec_desc";
                cbo.ValueMember = "rec_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        
        public int ToInt(String month)
        {
            try
            {
                return DateTime.ParseExact(month, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }

            return 0;
        }

        public String toAccountingFormat(Double amt)
        {
            try
            {
                return amt.ToString("#,##0.00;(#,##0.00);0.00");
            }
            catch (Exception er) { MessageBox.Show(er.Message);  }

            return "";
        }

        public Double toNormalDoubleFormat(String acct_amt)
        {
            Double amt = 0.00;

            try
            {
                if(String.IsNullOrEmpty(acct_amt))
                {
                    return 0.00;
                }
                else if (acct_amt.Contains("(") && acct_amt.Contains(")"))
                {
                    amt = Double.Parse(acct_amt, NumberStyles.AllowParentheses |
                                      NumberStyles.AllowThousands |
                                      NumberStyles.AllowDecimalPoint);
                }
                else
                    amt = Convert.ToDouble(acct_amt);
            }
            catch (Exception er) 
            { 
                //MessageBox.Show(er.Message); 
                amt = 0.00;
            }
            
            return amt;
        }


        //inventory
        public void load_vat(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("vat", "vat_code, vat_desc", "", "");
                cbo.DataSource = dt;
                cbo.DisplayMember = "vat_desc";
                cbo.ValueMember = "vat_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_saleunit(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("itmunit", "unit_id, unit_desc", "", " ORDER BY unit_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "unit_desc";
                cbo.ValueMember = "unit_id";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_unit_with_desc(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("itmunit", "unit_id, unit_shortcode ||' - '|| unit_desc", "", " ORDER BY unit_shortcode ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "unit_desc";
                cbo.ValueMember = "unit_id";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_account_title(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "", " ORDER BY at_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_account_title_code_asc(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m04", "at_code, at_desc", "", " ORDER BY at_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "at_desc";
                cbo.ValueMember = "at_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_stocklocation(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.get_location_list();
                cbo.DataSource = dt;
                cbo.DisplayMember = "whs_desc";
                cbo.ValueMember = "whs_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_supplier(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m07", "c_code, c_name", "", " ORDER BY c_name ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "c_name";
                cbo.ValueMember = "c_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_modeofpayment(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m10", "mp_code, mp_desc", "", " ORDER BY mp_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "mp_desc";
                cbo.ValueMember = "mp_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_category_asc_desc(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("itmgrp", "item_grp, grp_desc", "", " ORDER BY grp_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "grp_desc";
                cbo.ValueMember = "item_grp";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_category(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("itmgrp", "item_grp, grp_desc", "", " ORDER BY item_grp ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "grp_desc";
                cbo.ValueMember = "item_grp";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_generic(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("generic", "gen_code, gen_name", "cancel IS NULL OR cancel!='Y'", " ORDER BY gen_name ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "gen_name";
                cbo.ValueMember = "gen_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_brand(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("brand", "brd_code, brd_name", "cancel IS NULL OR cancel!='Y'", " ORDER BY brd_name ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "brd_name";
                cbo.ValueMember = "brd_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        /*

        public void load_terms(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m10", "mp_code, mp_desc", "isterms='Y' AND (cancel IS NULL OR cancel!='Y')", " ORDER BY mp_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "mp_desc";
                cbo.ValueMember = "mp_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_vat(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("vat", "vat_code, vat_desc", "cancel IS NULL OR cancel!='Y'", "");
                cbo.DataSource = dt;
                cbo.DisplayMember = "vat_desc";
                cbo.ValueMember = "vat_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }*/
        public void load_salesclerk(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.get_salesrep_list_with_canc();
                cbo.DataSource = dt;
                cbo.DisplayMember = "rep_name";
                cbo.ValueMember = "rep_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        /*
        public static int ToInt(String month)
        {
            try
            {
                return DateTime.ParseExact(month, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }

            return 0;
        } */

        public void load_item_asc_desc(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("items", "item_code, item_desc", "", " ORDER BY item_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "item_desc";
                cbo.ValueMember = "item_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_item(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("items", "item_code, item_desc", "", " ORDER BY item_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "item_desc";
                cbo.ValueMember = "item_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_charge(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("charge", "chg_code, chg_desc", "", " ORDER BY chg_type ASC, chg_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "chg_desc";
                cbo.ValueMember = "chg_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_paymenttype(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("m10", "mp_code, mp_desc", "", " ORDER BY mp_code ASC, mp_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "mp_desc";
                cbo.ValueMember = "mp_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_outlet(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("outlet", "out_code, out_desc", "", " ORDER BY out_code ASC, out_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "out_desc";
                cbo.ValueMember = "out_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_orhdr_inv_not_jrnlz(ComboBox cbo, String out_code)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("orhdr", "ord_code", " jrnlz IS NULL OR jrnlz!='Y' ", " ORDER BY ord_code ASC, out_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "ord_code";
                cbo.ValueMember = "ord_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_pinvhd_inv_not_jrnlz(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("pinvhd", "inv_num", " jrnlz IS NULL OR jrnlz!='Y' ", " ORDER BY inv_num ASC, out_desc ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "inv_num";
                cbo.ValueMember = "inv_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_pr_code(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("prhdr", "pr_code AS pr_code, pr_code ||' - '||reference ||' - '|| pr_date AS pr_desc", " cancel!='Y' ", " ORDER BY pr_code ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "pr_desc";
                cbo.ValueMember = "pr_code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_po_number(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("purhdr", "purc_ord AS pcode, purc_ord ||' - '|| supl_name AS cname", "", " ORDER BY purc_ord ASC");
                
                //dt = db.QueryBySQLCode("SELECT p.purc_ord AS pcode, p.purc_ord ||' - '|| m7.c_name AS cname FROM " + db.get_schema() + ".purhdr p LEFT JOIN " + db.get_schema() + ".m07 m7 ON p.supl_name=m7.c_code ORDER BY p.purc_ord");
                cbo.DataSource = dt;
                cbo.DisplayMember = "cname";
                cbo.ValueMember = "pcode";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_rr_number(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("rechdr", "rec_num AS rec_num, _reference", "trn_type='P' AND cancel !='Y'", "ORDER BY rec_num"); 
                cbo.DataSource = dt;
                cbo.DisplayMember = "_reference";
                cbo.ValueMember = "rec_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_dr_number(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("pinvhd", "inv_num, reference", "cancel != 'Y'", " ORDER BY inv_num ASC");
                cbo.DataSource = dt;
                cbo.DisplayMember = "reference";
                cbo.ValueMember = "inv_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_si_number(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("rechdr", "rec_num AS rec_num, _reference", "trn_type='I' AND cancel !='I'", "ORDER BY rec_num");
                cbo.DataSource = dt;
                cbo.DisplayMember = "_reference";
                cbo.ValueMember = "rec_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_tra_number(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("rechdr", "rec_num AS rec_num, _reference", "trn_type='T' AND cancel !='Y'", "ORDER BY rec_num");
                cbo.DataSource = dt;
                cbo.DisplayMember = "_reference";
                cbo.ValueMember = "rec_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_adj_number(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryOnTableWithParams("rechdr", "rec_num AS rec_num, _reference", "trn_type='A' AND cancel !='Y'", "ORDER BY rec_num");
                cbo.DataSource = dt;
                cbo.DisplayMember = "_reference";
                cbo.ValueMember = "rec_num";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        //HRIS


        public void load_dept(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_department WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DataSource = dt;
                cbo.DisplayMember = "dept_name";
                cbo.ValueMember = "deptid";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_section(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_depsection WHERE COALESCE(cancel,cancel,'')<>'Y'");
                if (dt.Rows.Count > 0)
                {
                    cbo.DataSource = dt;
                    cbo.DisplayMember = "section_name";
                    cbo.ValueMember = "secid";
                    cbo.SelectedIndex = -1;
                }
            }
            catch (Exception) { }
        }

        public void load_section(ComboBox cbo , String deptid)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_depsection WHERE deptid = '" + deptid + "'  AND COALESCE(cancel,cancel,'')<>'Y'");
                if (dt.Rows.Count > 0)
                {
                    cbo.DataSource = dt;
                    cbo.DisplayMember = "section_name";
                    cbo.ValueMember = "secid";
                    cbo.SelectedIndex = -1;
                }
            }
            catch (Exception) { }
        }
        public void load_position(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_position WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DataSource = dt;
                cbo.DisplayMember = "position_name";
                cbo.ValueMember = "postid";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_civil_status(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable(); 
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_civil_status");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_employee(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("select empid,concat(firstname,' ',lastname) as name from rssys.hr_employee");
                cbo.DataSource = dt;
                cbo.DisplayMember = "name";
                cbo.ValueMember = "empid";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        public void load_payroll_period(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT pay_code,concat(to_char(date_from, 'mm/dd/yyyy'),' To ',to_char(date_to, 'mm/dd/yyyy')) as period FROM rssys.hr_payrollpariod WHERE COALESCE(cancel,cancel,'')<>'Y'");

                cbo.DataSource = dt;
                cbo.DisplayMember = "period";
                cbo.ValueMember = "pay_code";
                cbo.SelectedIndex = -1;
            }
            catch { }
        }


        public void load_emp_stat(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_emp_status");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "statcode";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_rate_type(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_rate_type");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "ratecode";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_wtax(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_wtax");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_days(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_days");
                cbo.DataSource = dt;
                cbo.DisplayMember = "dayname";
                cbo.ValueMember = "day";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_loantype(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_loan_type");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_other_earnings(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_other_earnings WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        public void load_other_deductions(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_other_deductions WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }


        public void load_leave_type(ComboBox cbo)
        {
            try
            {
                DataTable dt = new DataTable();
                thisDatabase db = new thisDatabase();

                dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_leave_type WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DataSource = dt;
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }

        /*
        public void load_shift_schedule(ComboBox cbo)
        {
            try
            {
                thisDatabase db = new thisDatabase();

                cbo.DataSource = db.QueryBySQLCode("SELECT *, am_timein||'-'||am_timeout||'AM,'||pm_timein||'-'||pm_timeout||'PM' AS description FROM rssys.hr_shift_schedule WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DisplayMember = "description";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;
            }
            catch (Exception) { }
        }
        */

        // 10/24
        private int get_cbo_dropdownwidth(ComboBox cbo)
        {
            int maxWidth = 0, temp = 0;
            foreach (var obj in cbo.Items)
            {
                temp = TextRenderer.MeasureText(obj.ToString(), cbo.Font).Width;
                if (temp > maxWidth)
                {
                    maxWidth = temp;
                }
            }
            return maxWidth;
        }


        public void load_sss_brackets(ComboBox cbo)
        {
            try
            {
                thisDatabase db = new thisDatabase();

                cbo.DataSource = db.QueryBySQLCode("SELECT code, code ||', '||bracket1 || ' - ' || bracket2 AS bracket from rssys.hr_sss WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DisplayMember = "bracket";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;

                cbo.DropDownWidth = get_cbo_dropdownwidth(cbo);
            }
            catch (Exception) { }
        }
        public void load_pagibig_brackets(ComboBox cbo)
        {
            try
            {
                thisDatabase db = new thisDatabase();

                cbo.DataSource = db.QueryBySQLCode("SELECT code, code ||', '||bracket1 || ' - ' || bracket2 AS bracket from rssys.hr_hdmf WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DisplayMember = "bracket";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;

                cbo.DropDownWidth = get_cbo_dropdownwidth(cbo);
            }
            catch (Exception) { }
        }
        public void load_philhealth_brackets(ComboBox cbo)
        {
            try
            {
                thisDatabase db = new thisDatabase();

                cbo.DataSource = db.QueryBySQLCode("SELECT code, code ||', '||bracket1 || ' - ' || bracket2 AS bracket from rssys.hr_philhealth WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DisplayMember = "bracket";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;

                cbo.DropDownWidth = get_cbo_dropdownwidth(cbo);
            }
            catch (Exception) { }
        }


        // 10/27
        public void load_shift_schedule(ComboBox cbo)
        {
            try
            {
                thisDatabase db = new thisDatabase();

                cbo.DataSource = db.QueryBySQLCode("SELECT code, code ||', '|| CASE WHEN am_timein<>'' AND pm_timein<>'' THEN am_timein ||' - '|| am_timeout ||'AM/'|| pm_timein ||' - '|| pm_timeout||'PM' WHEN am_timein<>'' THEN am_timein ||' - '|| am_timeout||'AM' ELSE pm_timein ||' - '|| pm_timeout||'PM' END  AS shift from rssys.hr_shift_schedule WHERE COALESCE(cancel,cancel,'')<>'Y'");
                cbo.DisplayMember = "shift";
                cbo.ValueMember = "code";
                cbo.SelectedIndex = -1;

                cbo.DropDownWidth = get_cbo_dropdownwidth(cbo);
            }
            catch (Exception) { }
        }

    }
}
