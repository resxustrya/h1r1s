using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Reflection;
using Npgsql;

namespace Human_Resource_Information_System
{
    public class thisDatabase
    {
        public static String driveloc = "\\\\RIGHTAPPS\\RightApps\\";
        public static String comp_folder = "Eastland";
        public static String lcl_db = "pms_eastland";
        public static String svr_pass = "Rightech777";


        //public static String driveloc = "\\\\lourence-PC\\RightApps\\"; //"R:\\"
        //public static String comp_folder = "RA_System";//"Headway"; "";
        //public static String lcl_db = "inv_balai_merge";//"inv_headway"; "";
        //public static String svr_pass = "hahahehe";//"Rightech777";     "hahahehe";
        public static String schema_static = "rssys";//"rssys"; "ideasoft";

        public static String servers = System.IO.File.ReadAllText(driveloc + comp_folder + "\\HRPublish\\localDatabase.txt");
         public String schema = schema_static;
         GlobalMethod gm;
         NpgsqlConnection conn;

        public thisDatabase ()
        {
            gm = new GlobalMethod();
            conn = new NpgsqlConnection("Server=" + servers + ";Port=5432;User Id=postgres;Password=" + svr_pass + ";Database=" + lcl_db + ";");
        }

        public void OpenConn()
        {
            CloseConn();

            try
            {
                conn.Open();
                //MessageBox.Show("Connection State " + conn.State.ToString());
            }
            catch (Exception er)
            {
                MessageBox.Show("Connection Exception : " + er.Message);
            }
        }

        public void CloseConn()
        {
            try
            {
                conn.Close();
            }
            catch (Exception)
            {
                //MessageBox.Show(er.Message);
            }
        }

        //return user id if success, otherwise null
        public String validate_login(String user, String pass, String comp)
        {
            try
            {
                String flag = null;

                this.OpenConn();

                String SQL = "Select uid FROM " + schema + ".X08 WHERE uid='" + user.ToUpper() + "' AND pwd='" + pass + "'";

                NpgsqlCommand command = new NpgsqlCommand(SQL, conn);

                NpgsqlDataReader dr = command.ExecuteReader();

                if (dr.Read())
                    flag = dr[0].ToString();

                this.CloseConn();


                return flag;
            }
            catch (Exception)
            { }

            return null;
        }

        public String getFullName()
        {
            thisDatabase db = new thisDatabase();
            DataTable dt = db.QueryOnTableWithParams("x08", "opr_name", "", "");

            String fullName = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    fullName = dt.Rows[i]["opr_name"].ToString();
                }
            }
            catch (Exception)
            {

            }

            return (fullName != null) ? fullName : "NONAME";
        }

        public Boolean InsertOnTable(String table, String column, String value)
        {
            Boolean flag = false;

            if (GlobalClass.branch != "001" && GlobalClass.DontSendToMain == false)
            {
                thisDatabase2 db2 = new thisDatabase2();
                db2.InsertOnTable(table, column, value);
            }

            try
            {
                this.OpenConn();

                string SQL = "INSERT INTO " + this.schema + "." + table + "(" + column + ") VALUES (" + value + ")";
               // MessageBox.Show(SQL);
                NpgsqlCommand command = new NpgsqlCommand(SQL, conn);

                Int32 rowsaffected = command.ExecuteNonQuery();

                this.CloseConn();

                flag = true;
            }
            catch (Exception er)
            {
                flag = false;
                MessageBox.Show(er.Message);
            }

            return flag;
        }

        public Boolean UpdateOnTable(String table, String col_upd, String cond)
        {
            Boolean flag = false;

            if (GlobalClass.branch != "001" && GlobalClass.DontSendToMain == false)
            {
                thisDatabase2 db2 = new thisDatabase2();
                db2.UpdateOnTable(table, col_upd, cond);
            }

            try
            {
                this.OpenConn();

                if (cond != "")
                {
                    cond = " WHERE " + cond + "";
                }

                string SQL = "UPDATE " + this.schema + "." + table + " SET " + col_upd + "" + cond + ";";
               // MessageBox.Show(SQL);
                NpgsqlCommand command = new NpgsqlCommand(SQL, conn);

                Int32 rowsaffected = command.ExecuteNonQuery();

                this.CloseConn();

                flag = true;
            }
            catch (Exception er)
            {
                flag = false;
                MessageBox.Show(er.Message);
            }

            return flag;
        }

        public Boolean DeleteOnTable(String table, String cond)
        {
            Boolean flag = false;

            if (GlobalClass.branch != "001" && GlobalClass.DontSendToMain == false)
            {
                thisDatabase2 db2 = new thisDatabase2();
                db2.DeleteOnTable(table, cond);
            }
            
            try
            {
                this.OpenConn();

                string SQL = "DELETE FROM " + this.schema + "." + table + " WHERE " + cond + ";";
                //MessageBox.Show(SQL);
                NpgsqlCommand command = new NpgsqlCommand(SQL, conn);

                Int32 rowsaffected = command.ExecuteNonQuery();

                this.CloseConn();

                flag = true;
            }
            catch (Exception)
            { flag = false; }

            return flag;
        }

        /// EXTRA METHODS

        public Boolean isExists(String table, String sqlcheck)
        {
            DataTable dt;
            Boolean flag = false;

            try
            {
                dt = this.QueryOnTableWithParams(table, "1", sqlcheck, "");

                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception)
            { }

            return flag;
        }

        public String str_E(String str)
        {
            if (String.IsNullOrEmpty(str))
            {
                str = "'" + str + "'";
            }
            else
            {
                str = "$$" + str + "$$";
            }

            return str;
        }

        public String castToInteger(String col)
        {
            return "COALESCE(CAST(SUBSTRING(" + col + " FROM '([0-9]{1,10})') AS INTEGER), 0)";
            //            return "COALESCE(CAST(SUBSTRING(" + col + " FROM '([0-9]{1,10})') AS INTEGER), 0)";
        }

        /// END OF EXTRA METHODS 

        public DataTable QueryAllOnTable(string table)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            try
            {
                this.OpenConn();

                string SQL = "SELECT * FROM " + this.schema + "." + table + ";";

                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQL, conn);

                ds.Reset();

                da.Fill(ds);

                this.CloseConn();

                return ds.Tables[0];
            }
            catch (Exception er)
            {
                //MessageBox.Show(er.Message);
                return null;
            }
        }

        public DataTable QueryOnTableWithParams(string table, String param, String cond, String addcode)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            try
            {
                this.OpenConn();

                if (String.IsNullOrEmpty(cond))
                {
                    cond = "";
                }
                else
                {
                    cond = " WHERE " + cond;
                }

                string SQL = "SELECT " + param + " FROM " + this.schema + "." + table + " " + cond + " " + addcode;
                //MessageBox.Show(SQL);
                Console.WriteLine(SQL);
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQL, conn);

                ds.Reset();

                da.Fill(ds);

                this.CloseConn();

                return ds.Tables[0];
            }

            catch (Exception er)
            {
                //MessageBox.Show(er.Message);
                return null;
            }
        }

        public DataTable QueryBySQLCode(String SQL)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            try
            {
                this.OpenConn();

                //MessageBox.Show(SQL);
                Console.WriteLine(SQL);

                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQL, conn);

                ds.Reset();

                da.Fill(ds);

                this.CloseConn();

                return ds.Tables[0];
            }
            catch (Exception er)
            {
                //MessageBox.Show(er.Message);
                return null;
            }
        }

        public String get_m99comp_name()
        {
            String val = "";

            DataTable dt = this.QueryOnTableWithParams("m99", "comp_name", "", "");
            try
            {
                for (Int32 i = 0; i < dt.Rows.Count; i++)
                {
                    //tdate = Convert.ToDateTime(dt.Rows[i]["trnx_date"]).ToString("MMM dd, yyyy");
                    val = dt.Rows[i]["comp_name"].ToString();
                }
            }
            catch (Exception er)
            { }

            return val;
        }

        public String get_m99comp_addr()
        {
            String val = "";

            DataTable dt = this.QueryOnTableWithParams("m99", "comp_addr", "", "");

            for (Int32 i = 0; i < dt.Rows.Count; i++)
            {
                //tdate = Convert.ToDateTime(dt.Rows[i]["trnx_date"]).ToString("MMM dd, yyyy");
                val = dt.Rows[i]["comp_addr"].ToString();
            }

            return val;
        }

        //current default branch
        public String get_m99branch()
        {
            String val = "";

            DataTable dt = this.QueryOnTableWithParams("m99", "branch", "", "");
            try
            {
                for (Int32 i = 0; i < dt.Rows.Count; i++)
                {
                    //tdate = Convert.ToDateTime(dt.Rows[i]["trnx_date"]).ToString("MMM dd, yyyy");
                    val = dt.Rows[i]["branch"].ToString();
                }
            }
            catch (Exception er)
            { }

            return val;
        }

        public String get_systemdate(String format)
        {
            String tdate = "";
            DataTable dt = this.QueryOnTableWithParams("m99", "trnx_date", "", "");

            try
            {
                if (String.IsNullOrEmpty(format) == false)
                {
                    format = "yyyy-MM-dd";
                }

                if (GlobalClass.projcompany == "1" || GlobalClass.projcompany == "4")
                {
                    tdate = DateTime.Now.ToString(format);
                }
                else
                {
                    tdate = Convert.ToDateTime(dt.Rows[0]["trnx_date"]).ToString(format);
                }

                //temporary
                tdate = DateTime.Now.ToString(format);
            }
            catch (Exception er)
            {
                tdate = DateTime.Now.ToString("yyyy-MM-dd");

                MessageBox.Show("Error on date. Date/Time set to your workstation date and time. " + tdate + "\n\n Error :" + er.Message);
            }

            return tdate;
        }

        public String get_systemtime()
        {
            return DateTime.Now.ToString("HH:mm");
        }

        public String get_system_loc()
        {
            String system_loc = "";
            DateTime ldt;
            DataTable dt = this.QueryOnTableWithParams("m99", "system_loc", "", "");

            try
            {
                system_loc = dt.Rows[0]["system_loc"].ToString();
            }
            catch (Exception) { }

            return system_loc;
        }

        //get pk value that used for the table. input: m99col = column of m99
        public String get_pk(String m99col)
        {
            String pk = "";

            try
            {

                DataTable dt = this.QueryOnTableWithParams("m99", m99col, "", "");

                for (Int32 i = 0; i < dt.Rows.Count; i++)
                {
                    pk = dt.Rows[i][m99col].ToString();
                }
            }

            catch (Exception ex) { MessageBox.Show(ex.Message); }

            return pk;
        }

        public String get_nextincrementlimitchar(String val, int limit)
        {
            string newvalue = "";
            int len = val.Length;            
            var split = val.ToCharArray();
            int num = 0;
            
            if (split[0] == 'A')
            {
                split[0] = '0';
                var temp = new String(split);
                num = Convert.ToInt32(temp);
                num++;
            }
            else
            {                
                num = Convert.ToInt32(val);
                num++;
            }

            while (len > 0)
            {
                newvalue += "0";
                len--;
            }
            
            newvalue = num.ToString(newvalue);

            return newvalue;
        }

        public Boolean set_all_pk(String table, String pkcol, String pk_val, String cond, int limit)
        {
            String newpk = get_nextincrementlimitchar(pk_val, limit);

            return this.UpdateOnTable(table, pkcol + "='" + newpk + "'", cond);
        }

        //set new pk to m99.. the next increment value for the next input
        public Boolean set_pkincrement(String m99col, String val)
        {
            string newvalue = string.Empty;
            int len = val.Length;
            string split = val.Substring(2, len - 2);
            int num = Convert.ToInt32(split);

            num++;
            newvalue = val.Substring(0, 2) + num.ToString("0000");

            return this.UpdateOnTable("m99", m99col + "='" + newvalue + "'", "");
        }

        public Boolean set_pkm99(String m99col, String val)
        {
            return this.UpdateOnTable("m99", m99col + "='" + val + "'", "");
        }

        public Boolean set_cancel(String table, String code_or_cond)
        {
            return this.UpdateOnTable(table, "cancel='Y'", code_or_cond);
        }

        // get next increment value with 6 characters.. 
        public String get_nextincrement(String val)
        {
            string newvalue = string.Empty;

            int len = val.Length;
            //string split = val.Substring(2, len - 2);
            int num = Convert.ToInt32(val);
            num++;
            //newvalue = val.Substring(0, 2) + num.ToString("0000");

            newvalue = num.ToString("000000");

            return newvalue;
        }

        //get next increment value with limit of characters.. min: 3 chars. max: 8chars
        //public String get_nextincrementlimitchar(String val, int limit)
        //{
        //    string newvalue = string.Empty;
        //    //MessageBox.Show(val.ToString());
        //    int len = val.Length;
        //    //string split = jid.Substring(2, len - 2);
        //    int num = Convert.ToInt32(val);
        //    num++;
        //    //newvalue = jid.Substring(0, 2) + num.ToString("0000");

        //    if (limit == 3)
        //        newvalue = num.ToString("000");
        //    else if (limit == 4)
        //        newvalue = num.ToString("0000");
        //    else if (limit == 5)
        //        newvalue = num.ToString("00000");
        //    else if (limit == 6)
        //        newvalue = num.ToString("000000");
        //    else if (limit == 7)
        //        newvalue = num.ToString("0000000");
        //    else if (limit == 8)
        //        newvalue = num.ToString("00000000");

        //    return newvalue;
        //}

        //get the last value of key in a table.
        public String get_colval(String table, String col, String cond)
        {
            String pk = "";

            DataTable dt = this.QueryOnTableWithParams(table, col, cond, "ORDER BY " + col + " ASC");

            for (Int32 i = 0; i < dt.Rows.Count; i++)
            {
                pk = dt.Rows[i][col].ToString();
            }

            return pk;
        }

        public int get_cntrow(String table, String col, String cond)
        {
            DataTable dt = this.QueryOnTableWithParams(table, col, cond, "ORDER BY " + col + " ASC");

            return dt.Rows.Count;
        }

        public String get_schema()
        {
            return this.schema;
        }

        public DataTable get_alluserid()
        {
            return this.QueryOnTableWithParams("x08", "uid", "", "ORDER BY uid ASC");
        }

        public DataTable get_grouprights()
        {
            return this.QueryOnTableWithParams("x07", "*", "", "ORDER BY grp_id ASC");
        }

        public DataTable get_mainaccountlist()
        {
            return this.QueryOnTableWithParams("m01", "*", "", " ORDER BY mag_code ASC");
        }

        public DataTable get_subaccountlist()
        {
            return this.QueryOnTableWithParams("m02", "*", "", " ORDER BY cmp_code ASC");
        }

        public DataTable get_accountgrouplist()
        {
            return this.QueryOnTableWithParams("m03", "*", "", " ORDER BY acc_code ASC");
        }

        public DataTable get_accounttitlelist()
        {
            return this.QueryOnTableWithParams("m04", "*", "", " ORDER BY at_code ASC");
        }

        public DataTable get_journallist()
        {
            return this.QueryAllOnTable("m05");
        }

        public DataTable get_customerlist()
        {
            return this.QueryAllOnTable("m06");
        }

        public DataTable get_guestlist()
        {
            return this.QueryAllOnTable("guest");
        }

        public DataTable get_supplierlist()
        {
            return this.QueryAllOnTable("m07");
        }

        public DataTable get_costcenterlist()
        {
            return this.QueryAllOnTable("m08");
        }

        public DataTable get_moplist()
        {
            return this.QueryAllOnTable("m10");
        }
        public DataTable get_subcostcenterlist()
        {
            return this.QueryAllOnTable("subctr");
        }

        public DataTable get_location_list()
        {
            return this.QueryOnTableWithParams("whouse", "*", "", "ORDER BY whs_code ASC");
        }

        public DataTable get_itemgrp_list()
        {
            return this.QueryOnTableWithParams("itmgrp", "*", "", "ORDER BY item_grp ASC");
        }

        //status: O - open periods only, C - closed periods only or none ("") - all 
        public DataTable get_accountingperiodlist(String status)
        {
            String WHERE = "";

            if (status == "O")
            {
                WHERE = "closed='' OR closed is null";
            }
            else if (status == "C")
            {
                WHERE = "closed='Y'";
            }

            return this.QueryOnTableWithParams("x03", "*", WHERE, "ORDER BY fy ASC, mo ASC");
        }

        //returns D if Debit, C if Credit, otherwise empty string
        public String get_dr_cr_typ(String acc_code)
        {
            String dr_cr = "";
            try
            {
                DataTable dt = this.QueryOnTableWithParams("m03", "dr_cr", "acc_code='" + acc_code + "'", "");

                dr_cr = dt.Rows[0][0].ToString();
            }
            catch (Exception)
            {
                return "";
            }

            return dr_cr;
        }

        public String get_dr_cr_typof_at_code(String at_code)
        {
            String dr_cr = "";
            try
            {
                DataTable dt = this.QueryOnTableWithParams("m04", "dr_cr", "at_code='" + at_code + "'", "");

                dr_cr = dt.Rows[0][0].ToString();
            }
            catch (Exception)
            {
                return "";
            }

            return dr_cr;
        }

        public String get_bs_pl(String acc_code)
        {
            String mag_code = "";
            String bs_pl = "";
            try
            {
                DataTable dt = this.QueryBySQLCode("SELECT m02.mag_code FROM " + schema + ".m03  RIGHT JOIN " + schema + ".m02 ON m03.cmp_code=m02.cmp_code WHERE m03.acc_code='" + acc_code + "'");

                mag_code = dt.Rows[0][0].ToString();

                if (mag_code.StartsWith("1") || mag_code.StartsWith("2") || mag_code.StartsWith("3"))
                {
                    bs_pl = "B";
                }
                else
                {
                    bs_pl = "P";
                }
            }
            catch (Exception)
            {
                return "";
            }

            return bs_pl;
        }

        public DataTable get_journalentrylist(String fy, String mo, String jrnl, Boolean include_cancel)
        {
            String WHERE = "";
            if (include_cancel == false)
            {
                WHERE = " AND cancel IS null";
            }
            return this.QueryOnTableWithParams("tr01", "j_num AS \"Ref_Num\", t_date AS \"Date\", t_desc AS \"Description\", payee AS \"Paid To\", ck_num AS \"Check Number\", ck_date AS \"Check Date\", cancel AS \"Cancel\", user_id AS \"User ID\"", "j_code='" + jrnl + "' AND mo=" + mo + " AND fy=" + fy + "" + WHERE, "ORDER BY j_num ASC");
        }

        public DataTable get_journalentry_info(String fy, String mo, String j_code, String j_num)
        {
            //return this.QueryOnTableWithParams("tr01", "*", "j_code='" + j_code + "' AND mo=" + mo + " AND fy=" + fy + " AND j_num='"+ j_num +"'", "");

            return this.QueryBySQLCode("SELECT t1.*, t3.j_memo FROM " + schema + ".tr01 t1 LEFT JOIN " + schema + ".tr03 t3 ON t1.j_code=t3.j_code AND t1.j_num=t3.j_num WHERE t1.j_code='" + j_code + "' AND t1.mo=" + mo + " AND t1.fy=" + fy + " AND t1.j_num='" + j_num + "'");
        }

        public DataTable get_journalentry_info_accntlist(String j_code, String j_num)
        {
            return this.QueryBySQLCode("SELECT t2.seq_num, act.at_desc, t2.debit, t2.credit, t2.invoice, t2.sl_name, cc.cc_desc, t2.pay_code, t2.rep_code, t2.seq_desc, t2.at_code, t2.sl_code, t2.cc_code FROM " + schema + ".tr02 t2 LEFT JOIN " + schema + ".m04 act ON act.at_code=t2.at_code LEFT JOIN " + schema + ".m08 cc ON cc.cc_code=t2.cc_code WHERE j_code='" + j_code + "' AND j_num='" + j_num + "' ORDER BY " + castToInteger("t2.seq_num"));
        }

        public String get_explanation(String j_code, String j_num)
        {
            String exp = "";
            try
            {
                DataTable dt = this.QueryOnTableWithParams("tr03", "j_memo", "j_code='" + j_code + "' AND j_num='" + j_num + "'", "");

                exp = dt.Rows[0][0].ToString();
            }
            catch (Exception)
            {
                return "";
            }

            return exp;
        }

        //typ: 1=> Sales 2=> Cost of Sale 3=> Direct Purchase 4=> Purachase Return 5=> Stock Transaction
        public DataTable get_z_jrnllist(String dt_frm, String dt_to, int typ)
        {
            return this.QueryOnTableWithParams("z_jrnl", "*", "t_date BETWEEN '" + dt_frm + "' AND '" + dt_to + "' AND typ=" + typ.ToString(), "ORDER BY t_date, t_time");
        }

        public DataTable get_z_upload_stklist(String dt_frm, String dt_to, String typ)
        {
            return this.QueryOnTableWithParams("z_upload_stk", "*", "t_date BETWEEN '" + dt_frm + "' AND '" + dt_to + "' AND typ='" + typ.ToString() + "'", "ORDER BY t_date, t_time");
        }

        //return last sequence no, if no last sequence no. found, then, it will return 1 to start 1.
        public int get_jrnl_lastseq_num(String j_code, String j_num)
        {
            int seq_num = 0;
            try
            {
                DataTable dt = this.QueryOnTableWithParams("tr02", "MAX(" + castToInteger("seq_num") + ")", "j_code='" + j_code + "' AND j_num='" + j_num + "'", "");

                seq_num = Convert.ToInt32(dt.Rows[0][0].ToString()) + 1;
            }
            catch (Exception)
            {
                return 1;
            }

            return seq_num;
        }

        public String get_at_desc(String at_code)
        {
            String at_desc = "";
            try
            {
                DataTable dt = this.QueryOnTableWithParams("m04", "at_desc", "at_code='" + at_code + "'", "");

                at_desc = dt.Rows[0][0].ToString();
            }
            catch (Exception)
            {
                return "";
            }

            return at_desc;
        }

        public Double get_item_fcp(String itemcode)
        {
            Double fcp = 0.00;
            try
            {
                DataTable dt = this.QueryOnTableWithParams("items", "fcp", "item_code='" + itemcode + "'", "");

                fcp = Convert.ToDouble(dt.Rows[0][0].ToString());
            }
            catch (Exception)
            {
                return 0.00;
            }

            return fcp;
        }

        public Double get_item_fcp_compute(String itemcode)
        {
            Double fcp = 0.00;

            try
            {
                DataTable dt = QueryBySQLCode("SELECT DISTINCT F1.price FROM " + schema + ".stkcrd F1 WHERE F1.price = (SELECT price FROM (SELECT price, COUNT(*) AS count FROM " + schema + ".stkcrd F2 WHERE F2.item_code = '" + itemcode + "' AND trn_type='P' GROUP BY F2.price ORDER BY count DESC, price DESC) AS F5 LIMIT 1)");

                fcp = Convert.ToDouble(dt.Rows[0][0].ToString());
            }
            catch (Exception)
            {
                return 0.00;
            }

            return fcp;
        }

        public Double get_item_ave_cost_compute(String itemcode)
        {
            //String table = "reclne", col = "price", cond = "item_code='" + itemcode + "'";
            String table2 = "pinvln", col2 = "price", cond2 = "item_code='" + itemcode + "'";
            Double ave_cost = 0.00;
            Double sum = 0.00;
            Double count = 0.00;

            try
            {
                DataTable dt = QueryBySQLCode("SELECT SUM(rl.price) FROM " + schema + ".rechdr r RIGHT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num WHERE rl.item_code='" + itemcode + "' AND (r.trn_type!='I' OR r.trn_type!='A' OR r.trn_type!='T')");
                sum = gm.toNormalDoubleFormat(dt.Rows[0][0].ToString());
            }
            catch (Exception){ }

            try
            {
                DataTable dt = QueryBySQLCode("SELECT COUNT(rl.price) FROM " + schema + ".rechdr r RIGHT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num WHERE rl.item_code='" + itemcode + "' AND (r.trn_type!='I' OR r.trn_type!='A' OR r.trn_type!='T')");
                count = gm.toNormalDoubleFormat(dt.Rows[0][0].ToString());
            }
            catch (Exception) { }

            try
            {
                sum = sum + get_sum(table2, col2, cond2);

                count = count + get_count(table2, col2, cond2);

                ave_cost = sum / count;
            }
            catch (Exception)
            {
                return 0.00;
            }

            return ave_cost;
        }

        public Double get_avg(String table, String col_ave, String cond)
        {
            Double avg = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams(table, "AVE(" + col_ave + ")", cond, "");

                avg = gm.toNormalDoubleFormat(dt.Rows[0][0].ToString());
            }
            catch (Exception)
            {
                return 0.00;
            }

            return avg;
        }

        public Double get_sum(String table, String col_sum, String cond)
        {
            Double avg = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams(table, "SUM(" + col_sum + ")", cond, "");

                avg = gm.toNormalDoubleFormat(dt.Rows[0][0].ToString());
            }
            catch (Exception)
            {
                return 0.00;
            }

            return avg;
        }

        public Double get_count(String table, String col_cnt, String cond)
        {
            Double avg = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams(table, "COUNT(" + col_cnt + ")", cond, "");

                avg = gm.toNormalDoubleFormat(dt.Rows[0][0].ToString());
            }
            catch (Exception)
            {
                return 0.00;
            }

            return avg;
        }

        public Boolean has_subledger(String at_code)
        {
            Boolean has_sl = false;
            DataTable dt = new DataTable();

            if (at_code.Length <= 9)
            {
                dt = this.QueryOnTableWithParams("m04", "1", "at_code='" + at_code + "' AND sl='Y'", "");

                try
                {
                    if (dt.Rows[0][0].ToString() == "1")
                    {
                        has_sl = true;
                    }
                }
                catch (Exception) { }
            }
            return has_sl;

        }

        public Boolean is_payment_acct(String at_code)
        {
            Boolean flag = false;
            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".m04 WHERE at_code='" + at_code + "' AND acc_code='1000'");
            //payment='Y';");
            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception) { }

            return flag;
        }


        public Boolean is_check_acct(String at_code)
        {
            Boolean flag = false;
            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".m04 WHERE at_code='" + at_code + "' AND cib_acct='Y';");
            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception) { }

            return flag;
        }

        public Boolean is_liabilities(String at_code)
        {
            Boolean flag = false;
            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".m04 LEFT JOIN " + schema + ".m03 ON m04.acc_code=m03.acc_code LEFT JOIN " + schema + ".m02 ON m03.cmp_code=m02.cmp_code LEFT JOIN " + schema + ".m01 ON m02.mag_code=m01.mag_code WHERE m01.mag_code='20' AND m04.at_code='" + at_code + "'");

            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception) { }

            return flag;
        }

        public Boolean is_expenseORincome(String at_code)
        {
            Boolean is_ei = false;
            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".m04 LEFT JOIN " + schema + ".m03 ON m04.acc_code=m03.acc_code LEFT JOIN " + schema + ".m02 ON m03.cmp_code=m02.cmp_code LEFT JOIN " + schema + ".m01 ON m02.mag_code=m01.mag_code WHERE (m01.mag_code='40' OR m01.mag_code='60') AND m04.at_code='" + at_code + "'");

            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    is_ei = true;
                }
            }
            catch (Exception) { }

            return is_ei;
        }

        public DataTable get_unpaid_invoices(String j_code, String at_code, String subsidiary_code)
        {
            DataTable dt = null;
            try
            {
                /*
                if (j_code == "CDJ" || j_code == "CV" || j_code == "CV1") //only CDJ:  || j_code == "CV" || j_code == "CV1"
                {
                    //dt = this.QueryBySQLCode("SELECT t2.j_code, t2.j_num, t2.invoice, t1.t_date, SUM(t2.debit) AS DR, SUM(t2.credit) AS CR, t1.t_desc "
                    //                          + "FROM " + schema + ".tr01 t1 INNER JOIN " + schema + ".tr02 t2 ON (t1.j_code=t2.j_code AND t1.j_num=t2.j_num) "
                    //                          + "WHERE t2.sl_code='" + subsidiary_code + "' AND t2.at_code='" + at_code + "' AND (t2.j_code='CDJ' OR t2.j_code='PJ') "
                    //                          + "GROUP BY t2.invoice, t1.t_date, t1.t_desc, t2.j_code, t2.j_num "
                    //                          + "ORDER BY t2.invoice, t1.t_date ASC");
                    dt = this.QueryBySQLCode("SELECT t2.j_code, t2.j_num, t2.invoice, t1.t_date, SUM(t2.debit) AS DR, SUM(t2.credit) AS CR, t1.t_desc "
                                              + "FROM " + schema + ".tr01 t1 INNER JOIN " + schema + ".tr02 t2 ON (t1.j_code=t2.j_code AND t1.j_num=t2.j_num) "
                                              + "WHERE t2.sl_code='" + subsidiary_code + "' AND t2.at_code='" + at_code + "' AND (t2.j_code='CDJ' OR t2.j_code='PJ' OR t2.j_code='CV1' OR t2.j_code='CV') "
                                              + "GROUP BY t2.invoice, t1.t_date, t1.t_desc, t2.j_code, t2.j_num "
                                              + "ORDER BY t2.invoice, t1.t_date ASC");

                   

                    Console.Write("{0}", dt.Rows.Count);
                }
                else if (j_code == "CRJ" || j_code == "PJ" || j_code == "JV")
                {
                    dt = this.QueryBySQLCode("SELECT t2.j_code, t2.j_num, t2.invoice, t1.t_date, SUM(t2.debit) AS DR, SUM(t2.credit) AS CR, t1.t_desc "
                                              + "FROM " + schema + ".tr01 t1 INNER JOIN " + schema + ".tr02 t2 ON (t1.j_code=t2.j_code AND t1.j_num=t2.j_num) "
                                              + "WHERE t2.sl_code='" + subsidiary_code + "' AND t2.at_code='" + at_code + "' AND (t2.j_code='CRJ' OR t2.j_code='SJ' OR t2.j_code='JV' OR t2.j_code='PJ') "
                                              + "GROUP BY t2.invoice, t1.t_date, t1.t_desc, t2.j_code, t2.j_num "
                                              + "ORDER BY t2.invoice, t1.t_date ASC");
                } */

                dt = this.QueryBySQLCode("SELECT t2.j_code, t2.j_num, t2.invoice, t1.t_date, SUM(t2.debit) AS DR, SUM(t2.credit) AS CR, t1.t_desc "
                                             + "FROM " + schema + ".tr01 t1 INNER JOIN " + schema + ".tr02 t2 ON (t1.j_code=t2.j_code AND t1.j_num=t2.j_num) "
                                             + "WHERE t2.sl_code='" + subsidiary_code + "' AND t2.at_code='" + at_code + "'"
                                             + "GROUP BY t2.invoice, t1.t_date, t1.t_desc, t2.j_code, t2.j_num "
                                             + "ORDER BY t2.invoice, t1.t_date ASC");


            }
            catch (Exception) { }

            return dt;
        }

        public Double get_debit_amt_in_CDJ(String subsidiary_code, String inv)
        {
            Double amt = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams("tr06", "debit", "c_code='" + subsidiary_code + "' AND invoice='" + inv + "'", "");

                amt = Convert.ToDouble(dt.Rows[0][0].ToString());
            }
            catch (Exception) { }

            return amt;
        }

        public Double get_credit_amt_in_CDJ(String subsidiary_code, String inv)
        {
            Double amt = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams("tr06", "credit", "c_code='" + subsidiary_code + "' AND invoice='" + inv + "'", "");

                amt = Convert.ToDouble(dt.Rows[0][0].ToString());
            }
            catch (Exception) { }

            return amt;
        }

        public Double get_debit_amt_in_CRJ(String subsidiary_code, String inv)
        {
            Double amt = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams("tr05", "debit", "d_code='" + subsidiary_code + "' AND invoice='" + inv + "'", "");

                amt = Convert.ToDouble(dt.Rows[0][0].ToString());
            }
            catch (Exception) { }

            return amt;
        }

        public Double get_credit_amt_in_CRJ(String subsidiary_code, String inv)
        {
            Double amt = 0.00;

            try
            {
                DataTable dt = this.QueryOnTableWithParams("tr05", "credit", "d_code='" + subsidiary_code + "' AND invoice='" + inv + "'", "");

                amt = Convert.ToDouble(dt.Rows[0][0].ToString());
            }
            catch (Exception) { }

            return amt;
        }

        public Boolean upd_unpaid_invoices(String j_code, String subsidiary_code, String inv, Double dr, Double cr)
        {
            Boolean flag = false;

            try
            {
                if (j_code == "CDJ")
                {
                    flag = this.UpdateOnTable("tr06", "debit=" + dr.ToString("0.00") + ", credit=" + cr.ToString() + "", "c_code='" + subsidiary_code + "' AND invoice='" + inv + "'");
                }
                else if (j_code == "CRJ")
                {
                    flag = this.UpdateOnTable("tr05", "debit=" + dr.ToString("0.00") + ", credit=" + cr.ToString() + "", "d_code='" + subsidiary_code + "' AND invoice='" + inv + "'");
                }
            }
            catch (Exception) { }

            return flag;
        }

        public Boolean jrnlz_hotel(String j_code, String dtfrm, String dtto)
        {
            Boolean flag = false;
            String col = "", val = "";
            String fy = get_pk("fy"), mo = Convert.ToDateTime(get_systemdate("yyyy-MM-dd")).Month.ToString(), t_desc = "";
            String j_num = "", seq_num, at_code, sl_code, sl_name, cc_code = "", scc_code = "", invoice, seq_desc = "", rep_code = "", pay_code = "", chg_type = "C";
            String table = "", chg_code = "", chg_num = "", chg_date = "";
            int lnno = 1;
            DataTable dt = new DataTable();
            DateTime t_date = new DateTime();
            DateTime next_date = new DateTime();
            String dr_cr = "", has_sl = "", next_regnum = "";
            Double debit = 0.00, credit = 0.00, tot_dr = 0.00, tot_cr = 0.00, amount = 0.00;

            //fetch all data from chgfil
            dt = this.QueryBySQLCode("SELECT \'chgfil\' AS tbl,  cf.reg_num,  cf.chg_code, c.chg_type, cf.chg_num,  cf.rom_code,  reference,  amount,  chg_date,  c.at_code, gf.acct_no AS sl_code, gf.full_name AS sl_name, c.cc_code, c.scc_code, m4.dr_cr, m4.sl, g.at_code AS g_at_code, jrnlz FROM rssys.chgfil cf LEFT JOIN rssys.charge c ON c.chg_code=cf.chg_code LEFT JOIN rssys.gfolio gf ON gf.reg_num=cf.reg_num LEFT JOIN rssys.m04 m4 ON m4.at_code=c.at_code LEFT JOIN rssys.guest g ON g.acct_no=gf.acct_no WHERE cf.chg_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' AND (cf.jrnlz!='Y' OR cf.jrnlz IS NULL) UNION ALL SELECT \'chghist\' AS tbl, ch.reg_num,  ch.chg_code, c.chg_type, ch.chg_num,  ch.rom_code,  reference,  amount,  chg_date,  c.at_code, gh.acct_no AS sl_code, gh.full_name AS sl_name, c.cc_code, c.scc_code, m4.dr_cr, m4.sl, g.at_code AS g_at_code, jrnlz FROM rssys.chghist ch LEFT JOIN rssys.charge c ON c.chg_code=ch.chg_code LEFT JOIN rssys.gfhist gh ON gh.reg_num=ch.reg_num LEFT JOIN rssys.m04 m4 ON m4.at_code=c.at_code LEFT JOIN rssys.guest g ON g.acct_no=gh.acct_no WHERE ch.chg_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' AND (ch.jrnlz!='Y' OR ch.jrnlz IS NULL) ORDER BY chg_date, reg_num;");

            try
            {
                if (dt != null)
                {
                    //loop by date and by regnum
                    //
                    if (dt.Rows.Count > 0)
                    {
                        //next_date = Convert.ToDateTime(dt.Rows[0]["chg_date"].ToString()).AddDays(1);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            t_date = Convert.ToDateTime(dt.Rows[0]["chg_date"].ToString());
                            next_date = t_date;
                            lnno = 1;

                            j_num = get_colval("m05", "j_num", "j_code='" + j_code + "'");
                            t_desc = "HOTEL TRANSACTIONS " + t_date.ToString("MM/dd/yy");

                            if (add_jrnl(fy, mo, j_code, j_num, t_desc, "", "", null, t_date.ToString("yyyy-MM-dd"), DateTimePicker.MinimumDateTime.ToString("yyyy-MM-dd")))
                            {
                                UpdateOnTable("m05", "j_num='" + get_nextincrement(j_num) + "'", "j_code='" + j_code + "'");

                                while (t_date == next_date)
                                {
                                    table = dt.Rows[i]["tbl"].ToString();

                                    seq_num = lnno.ToString();
                                    invoice = dt.Rows[i]["reg_num"].ToString();
                                    at_code = dt.Rows[i]["at_code"].ToString();
                                    cc_code = dt.Rows[i]["cc_code"].ToString();
                                    scc_code = dt.Rows[i]["scc_code"].ToString();
                                    seq_desc = dt.Rows[i]["reference"].ToString();
                                    chg_code = dt.Rows[i]["chg_code"].ToString();
                                    chg_num = dt.Rows[i]["chg_num"].ToString();
                                    chg_date = dt.Rows[i]["chg_date"].ToString();
                                    chg_type = dt.Rows[i]["chg_type"].ToString();
                                    dr_cr = dt.Rows[i]["dr_cr"].ToString();
                                    amount = gm.toNormalDoubleFormat(dt.Rows[i]["amount"].ToString());
                                    debit = 0.00;
                                    credit = 0.00;
                                    sl_code = null;
                                    sl_name = null;

                                    if (chg_type == "P")
                                        amount = amount * -1;

                                    if (dr_cr == "D")
                                    {
                                        debit = amount;

                                        if (debit < 0)
                                        {
                                            credit = amount;
                                            debit = 0.00;
                                            tot_cr += credit;
                                        }
                                        else
                                        {
                                            tot_dr += debit;
                                        }
                                    }
                                    else
                                    {
                                        credit = amount;

                                        if (credit < 0)
                                        {
                                            debit = amount;
                                            credit = 0.00;
                                            tot_dr += debit;
                                        }
                                        else
                                        {
                                            tot_cr += credit;
                                        }
                                    }

                                    if (has_sl == "Y")
                                    {
                                        sl_code = dt.Rows[i]["sl_code"].ToString();
                                        sl_name = dt.Rows[i]["sl_name"].ToString();

                                        if (String.IsNullOrEmpty(sl_code) == false)
                                        {
                                            upd_unpaid_invoices(j_code, sl_code, invoice, debit, credit);
                                        }
                                    }
                                    add_jrnl_entry(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, null, null);

                                    flag = UpdateOnTable(table, "jrnlz='Y'", "chg_date='" + chg_date + "' AND chg_num='" + chg_num + "' AND chg_code='" + chg_code + "'");

                                    lnno++;
                                    next_regnum = invoice;

                                    if (dt.Rows.Count > i + 1)
                                    {
                                        next_regnum = dt.Rows[i + 1]["reg_num"].ToString();
                                    }

                                    if ((next_regnum != invoice) || (dt.Rows.Count == i + 1))
                                    {
                                        seq_num = lnno.ToString();
                                        invoice = dt.Rows[i]["reg_num"].ToString();
                                        at_code = dt.Rows[i]["g_at_code"].ToString();
                                        sl_code = dt.Rows[i]["sl_code"].ToString();
                                        sl_name = dt.Rows[i]["sl_name"].ToString();
                                        cc_code = "";
                                        scc_code = "";
                                        seq_desc = "";
                                        chg_code = "";
                                        chg_num = "";
                                        chg_date = "";
                                        debit = 0.00;
                                        credit = 0.00;

                                        if (tot_dr > tot_cr)
                                        {
                                            credit = tot_dr - tot_cr;
                                        }
                                        else
                                        {
                                            debit = tot_cr - tot_dr;
                                        }
                                        add_jrnl_entry(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, null, null);

                                        tot_dr = 0.00;
                                        tot_cr = 0.00;
                                        lnno++;
                                    }

                                    i++;
                                    //error on next date upon the end;
                                    try
                                    {
                                        next_date = Convert.ToDateTime(dt.Rows[i]["chg_date"].ToString());
                                    }
                                    catch (Exception) { next_date = t_date.AddDays(1); }
                                } // end while
                            } // end if
                        }//end for
                    }

                }
            }
            catch (Exception er) { MessageBox.Show("Error on journalizing. \n" + er.Message); }

            return flag;
        }


        public Boolean jrnlz_sales(String j_code, String out_code, String dtfrm, String dtto)
        {
            Boolean flag = false;
            String col = "", val = "";
            String fy = get_pk("fy"), mo = Convert.ToDateTime(get_systemdate("yyyy-MM-dd")).Month.ToString(), t_desc = "";
            String j_num = "", seq_num, at_code, sl_code, sl_name, cc_code = "", scc_code = "", invoice, seq_desc = "", rep_code = "", pay_code = "", chg_type = "C", or_code = "", or_lne = "", item_code = "", item_desc = "", s_unit = "", recv_qty = "", s_price = "", whs = "";
            String table = "", chg_code = "", chg_num = "", chg_date = "";
            int lnno = 1;
            DataTable dt = new DataTable();
            DateTime t_date = new DateTime();
            DateTime next_date = new DateTime();
            String dr_cr = "", has_sl = "", next_regnum = "";
            Double debit = 0.00, credit = 0.00, tot_dr = 0.00, tot_cr = 0.00, amount = 0.00;

            //fetch all data from chgfil
            dt = this.QueryBySQLCode("SELECT o.out_code, o.ord_code AS or_code, o.ord_date, o.pay_code, o.debt_code AS sl_code, o.debt_name AS sl_name, o.inv_num, o.whs_code, o.rep_code, 'Outlet:' || o.out_code ||', Ord.No.:'|| o.ord_code AS reference, ol.ln_num AS or_lne, ol.item_code, ol.item_desc, ol.unit, ol.ord_qty AS recv_qty, ol.price, ol.ln_amnt AS amount, ol.disc_amnt, ol.disc2, ig.acct_sales AS at_code FROM rssys.orhdr o LEFT JOIN rssys.orlne ol ON ol.ord_code=o.ord_code LEFT JOIN rssys.outlet otl ON o.out_code=otl.out_code LEFT JOIN rssys.items i ON i.item_code=ol.item_code LEFT JOIN rssys.itmgrp ig ON ig.item_grp=i.item_grp WHERE o.out_code='" + out_code + "' AND (o.jrnlz!='Y' OR o.jrnlz IS NULL) AND o.ord_date BETWEEN '" + dtfrm + "' AND '" + dtto + "'");

            //gfolio or gfhist
            try
            {
                if (dt != null)
                {
                    //loop by date and by regnum
                    //
                    if (dt.Rows.Count > 0)
                    {
                        //next_date = Convert.ToDateTime(dt.Rows[0]["chg_date"].ToString()).AddDays(1);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            t_date = Convert.ToDateTime(dt.Rows[0]["ord_date"].ToString());
                            next_date = t_date;
                            lnno = 1;

                            j_num = get_colval("m05", "j_num", "j_code='" + j_code + "'");
                            t_desc = "OUTLET SALES TRANSACTIONS " + t_date.ToString("MM/dd/yy");

                            if (add_jrnl(fy, mo, j_code, j_num, t_desc, "", "", null, t_date.ToString("yyyy-MM-dd"), DateTimePicker.MinimumDateTime.ToString("yyyy-MM-dd")))
                            {
                                UpdateOnTable("m05", "j_num='" + get_nextincrement(j_num) + "'", "j_code='" + j_code + "'");

                                while (t_date == next_date)
                                {
                                    table = "orhdr"; //dt.Rows[i]["tbl"].ToString();

                                    seq_num = lnno.ToString();
                                    invoice = dt.Rows[i]["ord_code"].ToString();
                                    at_code = dt.Rows[i]["at_code"].ToString();
                                    cc_code = null; // dt.Rows[i]["cc_code"].ToString();
                                    scc_code = null; // dt.Rows[i]["scc_code"].ToString();
                                    seq_desc = dt.Rows[i]["reference"].ToString();
                                    pay_code = dt.Rows[i]["pay_code"].ToString();
                                    rep_code = dt.Rows[i]["rep_code"].ToString();
                                    chg_code = null; //dt.Rows[i]["chg_code"].ToString();
                                    chg_num = null; ////dt.Rows[i]["chg_num"].ToString();
                                    chg_date = null; // dt.Rows[i]["chg_date"].ToString();
                                    chg_type = null; // dt.Rows[i]["chg_type"].ToString();
                                    dr_cr = null; //dt.Rows[i]["dr_cr"].ToString();
                                    amount = gm.toNormalDoubleFormat(dt.Rows[i]["amount"].ToString());
                                    debit = 0.00;
                                    credit = 0.00;
                                    sl_code = null;
                                    sl_name = null;

                                    if (pay_code != "DON")
                                    {
                                        at_code = get_at_code_payment(pay_code);
                                        sl_code = dt.Rows[i]["sl_code"].ToString();
                                        sl_name = dt.Rows[i]["sl_name"].ToString();
                                        debit = amount;
                                        add_jrnl_entry(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, null, null);

                                        at_code = dt.Rows[i]["at_code"].ToString();
                                        or_code = dt.Rows[i]["or_code"].ToString();
                                        or_lne = dt.Rows[i]["or_lne"].ToString();

                                        sl_code = null;
                                        sl_name = null;
                                        item_code = dt.Rows[i]["or_lne"].ToString();
                                        item_desc = dt.Rows[i]["or_lne"].ToString();
                                        s_unit = dt.Rows[i]["or_lne"].ToString();
                                        recv_qty = dt.Rows[i]["or_lne"].ToString();
                                        s_price = dt.Rows[i]["or_lne"].ToString();
                                        whs = dt.Rows[i]["or_lne"].ToString();

                                        debit = 0.00;
                                        credit = amount;

                                        //credit side sales
                                        add_jrnl_entry_with_item(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, or_code, or_lne, item_code, item_desc, s_unit, recv_qty, s_price, whs);
                                    }

                                    else
                                    {
                                        at_code = "";
                                        sl_code = dt.Rows[i]["sl_code"].ToString();
                                        sl_name = dt.Rows[i]["sl_name"].ToString();
                                        debit = amount;
                                        add_jrnl_entry(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, null, null);
                                    }

                                    if (has_sl == "Y")
                                    {
                                        if (String.IsNullOrEmpty(sl_code) == false)
                                        {
                                            upd_unpaid_invoices(j_code, sl_code, invoice, debit, credit);
                                        }
                                    }

                                    add_jrnl_entry(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, or_code, or_lne);

                                    flag = UpdateOnTable(table, "jrnlz='Y'", "chg_date='" + chg_date + "' AND chg_num='" + chg_num + "' AND chg_code='" + chg_code + "'");

                                    lnno++;
                                    next_regnum = invoice;

                                    if (dt.Rows.Count > i + 1)
                                    {
                                        next_regnum = dt.Rows[i + 1]["reg_num"].ToString();
                                    }

                                    if ((next_regnum != invoice) || (dt.Rows.Count == i + 1))
                                    {
                                        seq_num = lnno.ToString();
                                        invoice = dt.Rows[i]["reg_num"].ToString();
                                        at_code = dt.Rows[i]["g_at_code"].ToString();
                                        sl_code = dt.Rows[i]["sl_code"].ToString();
                                        sl_name = dt.Rows[i]["sl_name"].ToString();
                                        cc_code = "";
                                        scc_code = "";
                                        seq_desc = "";
                                        chg_code = "";
                                        chg_num = "";
                                        chg_date = "";
                                        debit = 0.00;
                                        credit = 0.00;

                                        if (tot_dr > tot_cr)
                                        {
                                            credit = tot_dr - tot_cr;
                                        }
                                        else
                                        {
                                            debit = tot_cr - tot_dr;
                                        }

                                        add_jrnl_entry(j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, null, debit.ToString("0.00"), credit.ToString("0.00"), invoice, seq_desc, rep_code, pay_code, null, null);

                                        tot_dr = 0.00;
                                        tot_cr = 0.00;
                                        lnno++;
                                    }

                                    i++;
                                    //error on next date upon the end;
                                    try
                                    {
                                        next_date = Convert.ToDateTime(dt.Rows[i]["chg_date"].ToString());
                                    }
                                    catch (Exception) { break; }
                                } // end while
                            } // end if
                        }//end for
                    }
                }
            }
            catch (Exception)
            { }


            return flag;
        }

        public String get_at_code_payment(String mp_code)
        {
            String at_code = "";

            if (mp_code == "CSH")
            {
                mp_code = "101";
            }
            else if (mp_code == "CRD")
            {
                mp_code = "102";
            }
            else if (mp_code == "CRD")
            {
                mp_code = "103";
            }
            else if (mp_code == "CRD")
            {
                mp_code = "104";
            }
            else if (mp_code == "CRD")
            {
                mp_code = "105";
            }
            else if (mp_code == "TRM")
            {
                mp_code = "200"; //15DAYS
            }
            else if (mp_code == "DON")
            {
                mp_code = "101";
            }

            DataTable dt = this.QueryOnTableWithParams("m10", "mp_desc", "mp_code='" + mp_code + "'", "");

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    at_code = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return at_code;
        }

        public Boolean jrnlz_purchases(String j_code, String dtfrm, String dtto)
        {
            Boolean flag = false;


            //add purchase return
            if (jrnlz_purcahseReturn(j_code, dtfrm, dtto))
            {
                flag = true;
            }

            return flag;
        }

        private Boolean jrnlz_purcahseReturn(String j_code, String dtfrm, String dtto)
        {
            Boolean flag = false;

            return flag;
        }

        public Boolean jrnlz_stockTransaction(String j_code, String dtfrm, String dtto)
        {
            Boolean flag = false;

            return flag;
        }


        public Boolean add_jrnl(String fy, String mo, String j_code, String j_num, String t_desc, String payee, String ck_num, String relsd, String t_date, String ck_date)
        {
            if (isExists("payee", "payee=" + str_E(payee) + "") == false && String.IsNullOrEmpty(payee) == false)
            {
                this.InsertOnTable("payee", "payee", "" + str_E(payee) + "");
            }

            return this.InsertOnTable("tr01", "fy, mo, j_code, j_num, t_date, t_desc, payee, ck_date, ck_num, user_id, sysdate, systime, relsd", "'" + fy + "', '" + mo + "', '" + j_code + "', '" + j_num + "', '" + t_date + "', " + str_E(t_desc) + ", " + str_E(payee) + ", '" + ck_date + "', '" + ck_num + "', '" + GlobalClass.username + "', '" + this.get_systemdate("") + "', '" + DateTime.Now.ToString("HH:mm") + "', '" + relsd + "'");
        }

        public Boolean add_jrnl_explanation(String j_code, String j_num, String expl)
        {
            return this.InsertOnTable("tr03", "j_code, j_num, j_memo", "'" + j_code + "', '" + j_num + "', " + str_E(expl) + "");
        }

        public Boolean add_jrnl_entry(String j_code, String j_num, String seq_num, String at_code, String sl_code, String sl_name, String cc_code, String prj_code, String debit, String credit, String invoice, String seq_desc, String rep_code, String pay_code, String or_code, String or_lne)
        {
            GlobalClass gc = new GlobalClass();

            if (String.IsNullOrEmpty(or_lne))
            {
                or_lne = "0";
            }

            debit = debit.Replace(",", "");
            credit = credit.Replace(",", "");
            debit = gc.toNormalDoubleFormat(debit).ToString("0.00");
            credit = gc.toNormalDoubleFormat(credit).ToString("0.00");

            return this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, prj_code, debit, credit, invoice, seq_desc, rep_code, pay_code, or_code, or_lne", "'" + j_code + "', '" + j_num + "', '" + seq_num + "', '" + at_code + "', '" + sl_code + "', " + str_E(sl_name) + ", '" + cc_code + "', '" + prj_code + "', '" + debit + "', '" + credit + "', '" + invoice + "', " + str_E(seq_desc) + ", '" + rep_code + "', '" + pay_code + "', " + str_E(or_code) + ", '" + or_lne + "'");
        }

        public Boolean add_jrnl_entry_with_item(String j_code, String j_num, String seq_num, String at_code, String sl_code, String sl_name, String cc_code, String prj_code, String debit, String credit, String invoice, String seq_desc, String rep_code, String pay_code, String or_code, String or_lne, String item_code, String item_desc, String unit, String recv_qty, String price, String whouse)
        {
            GlobalClass gc = new GlobalClass();

            if (String.IsNullOrEmpty(or_lne))
            {
                or_lne = "0";
            }

            debit = debit.Replace(",", "");
            credit = credit.Replace(",", "");
            debit = gc.toNormalDoubleFormat(debit).ToString("0.00");
            credit = gc.toNormalDoubleFormat(credit).ToString("0.00");


            return this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, sl_code, sl_name, cc_code, prj_code, debit, credit, invoice, seq_desc, rep_code, pay_code, or_code, or_lne, item_code,  item_desc, unit, recv_qty, price, whouse", "'" + j_code + "', '" + j_num + "', '" + seq_num + "', '" + at_code + "', '" + sl_code + "', " + str_E(sl_name) + ", '" + cc_code + "', '" + prj_code + "', '" + debit + "', '" + credit + "', '" + invoice + "', " + str_E(seq_desc) + ", '" + rep_code + "', '" + pay_code + "', " + str_E(or_code) + ", '" + or_lne + ", '" + item_code + ", '" + item_desc + ", '" + unit + ", '" + recv_qty + ", '" + price + ", '" + whouse + "'");
        }

        public Boolean add_jrnlz_jrnl(String fy, String mo, String j_code, String j_num, String t_desc, String ord_date, String t_date, String t_time)
        {
            return this.InsertOnTable("tr01", "fy, mo, j_code, j_num, t_date, t_desc, payee, user_id, sysdate, systime", "'" + fy + "', '" + mo + "', '" + j_code + "', '" + j_num + "', '" + ord_date + "', " + str_E(t_desc) + ", '', '" + GlobalClass.username + "', '" + t_date + "', '" + t_time + "'");
        }

        public Boolean add_jrnlz_jrnl_entry(String j_code, String j_num, String at_code, String sl_code, String sl_name, String amt, String invoice, String pay_code, Item_Array itm)
        {
            GlobalClass gc = new GlobalClass();
            Double vat_rate = Convert.ToDouble(this.get_pk("vat_rate")) / 100;
            Double tax_amt = Convert.ToDouble(amt) / (100 + vat_rate);
            Double net_sales = Convert.ToDouble(amt) - tax_amt;
            String mi_at_code = "1300"; // merchandise inventory Account Code
            int seq_num = 1;

            if (j_code == "SJ")
            {
                //debit - recievable or cash
                if (pay_code == "COD")
                {
                    at_code = "1102"; //CASH ON HAND Account Code

                    this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, debit, invoice",
                        "'" + j_code + "', '" + j_num + "', 1, '" + at_code + "', " + amt + ", '" + invoice + "'");
                }
                else
                {
                    this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, sl_code, sl_name, debit, invoice",
                        "'" + j_code + "', '" + j_num + "', 1, '" + at_code + "', '" + sl_code + "', " + str_E(sl_name) + ", " + amt + ", '" + invoice + "'");
                }

                //credit - sales account
                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, credit, invoice",
                    "'" + j_code + "', '" + j_num + "', 2, '4000', " + net_sales.ToString("0.00") + ", '" + invoice + "'");
                //credit -tax
                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, credit, invoice",
                    "'" + j_code + "', '" + j_num + "', 3, '2300', " + tax_amt.ToString("0.00") + ", '" + invoice + "'");
            }
            else if (j_code == "PJ")
            {
                seq_num = get_jrnl_lastseq_num(j_code, j_num);

                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, debit, invoice, item_code, item_desc, unit, recv_qty, price",
                    "'" + j_code + "', '" + j_num + "', " + seq_num.ToString() + ", '" + mi_at_code + "', " + amt + ", '" + invoice + "', '" + itm.item_code + "', " + str_E(itm.item_desc) + ", '" + itm.unit + "', '" + itm.recv_qty + "', " + itm.price);

                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, sl_code, sl_name, credit, invoice",
                    "'" + j_code + "', '" + j_num + "', " + (seq_num + 1).ToString() + ", '" + at_code + "', '" + sl_code + "', " + str_E(sl_name) + ", " + amt + ", '" + invoice + "'");
            }
            //return purchase
            else if (j_code == "-PJ")
            {
                j_code = "PJ";
                seq_num = get_jrnl_lastseq_num(j_code, j_num);

                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, sl_code, sl_name, debit, invoice",
                   "'" + j_code + "', '" + j_num + "', " + (seq_num + 1).ToString() + ", '" + at_code + "', '" + sl_code + "', " + str_E(sl_name) + ", " + amt + ", '" + invoice + "'");

                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, credit, invoice, item_code, item_desc, unit, recv_qty, price",
                    "'" + j_code + "', '" + j_num + "', " + seq_num.ToString() + ", '" + mi_at_code + "', " + amt + ", '" + invoice + "', '" + itm.item_code + "', " + str_E(itm.item_desc) + ", '" + itm.unit + "', '" + itm.recv_qty + "', " + itm.price);
            }

            return true;
        }

        //TO BE DEVELOP
        public Boolean add_jrnl_entry_PJ(String j_code, String j_num, String at_code, String sl_code, String sl_name, String amt, String invoice, String pay_code)
        {
            GlobalClass gc = new GlobalClass();
            Double vat_rate = Convert.ToDouble(this.get_pk("vat_rate")) / 100;
            Double tax_amt = Convert.ToDouble(amt) / (100 + vat_rate);
            Double net_sales = Convert.ToDouble(amt) - tax_amt;

            //debit - recievable or cash
            if (pay_code == "COD")
            {
                at_code = "1102";
                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, debit, invoice", "'" + j_code + "', '" + j_num + "', 1, '" + at_code + "', " + amt + ", '" + invoice + "'");
            }
            else
            {
                this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, sl_code, sl_name, debit, invoice", "'" + j_code + "', '" + j_num + "', 1, '" + at_code + "', '" + sl_code + "', " + str_E(sl_name) + ", " + amt + ", '" + invoice + "'");
            }

            //credit - sales account
            this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, credit, invoice", "'" + j_code + "', '" + j_num + "', 2, '4000', " + net_sales.ToString("0.00") + ", '" + invoice + "'");
            //credit -tax
            this.InsertOnTable("tr02", "j_code, j_num, seq_num, at_code, credit, invoice", "'" + j_code + "', '" + j_num + "', 3, '2300', " + tax_amt.ToString("0.00") + ", '" + invoice + "'");

            return true;
        }

        public Boolean upd_jrnl(String fy, String mo, String j_code, String j_num, String t_desc, String payee, String ck_num, String relsd, String t_date, String ck_date)
        {
            if (isExists("payee", "payee=" + str_E(payee) + "") == false)
            {
                this.InsertOnTable("payee", "payee", "" + str_E(payee) + "");
            }


            return this.UpdateOnTable("tr01", "fy = '" + fy + "', mo = '" + mo + "', j_code = '" + j_code + "', j_num = '" + j_num + "', t_date = '" + t_date + "', t_desc = " + str_E(t_desc) + ", payee = " + str_E(payee) + ", ck_date = '" + ck_date + "', ck_num ='" + ck_num + "', user_id = '" + GlobalClass.username + "', sysdate = '" + this.get_systemdate("") + "', systime = '" + DateTime.Now.ToString("HH:mm") + "', relsd = '" + relsd + "'", "j_code = '" + j_code + "' AND j_num = '" + j_num + "'");
        }

        public Boolean upd_jrnl_explanation(String j_code, String j_num, String expl)
        {
            return this.UpdateOnTable("tr03", "j_memo = " + str_E(expl), "j_code = '" + j_code + "' AND j_num = '" + j_num + "'");
        }

        public Boolean upd_jrnl_entry(String j_code, String j_num, String seq_num, String at_code, String sl_code, String sl_name, String cc_code, String prj_code, String debit, String credit, String invoice, String seq_desc, String rep_code, String pay_code, String or_code, String or_lne)
        {
            GlobalClass gc = new GlobalClass();

            if (String.IsNullOrEmpty(or_lne))
            {
                or_lne = "0";
            }

            debit = debit.Replace(",", "");
            credit = credit.Replace(",", "");
            debit = gc.toNormalDoubleFormat(debit).ToString("0.00");
            credit = gc.toNormalDoubleFormat(credit).ToString("0.00");

            return this.UpdateOnTable("tr02", "at_code = '" + at_code + "', sl_code = '" + sl_code + "', sl_name = " + str_E(sl_name) + ", cc_code = '" + cc_code + "', prj_code = '" + prj_code + "', debit = '" + debit + "', credit = '" + credit + "', invoice = '" + invoice + "', seq_desc = " + str_E(seq_desc) + ", rep_code = '" + rep_code + "', pay_code = '" + pay_code + "', or_code = '" + or_code + "', or_lne = '" + or_lne + "'", "j_code = '" + j_code + "' AND j_num = '" + j_num + "' AND seq_num='" + seq_num + "'");
        }

        public Boolean del_jrnl(String j_code, String j_num, String t_desc)
        {
            return this.UpdateOnTable("tr01", "t_desc='CANCELLED-" + t_desc + "', cancel='Y'", "j_code = '" + j_code + "' AND j_num = '" + j_num + "'");
        }

        public Boolean del_jrnl_explanation(String j_code, String j_num)
        {
            return this.DeleteOnTable("tr03", "j_code = '" + j_code + "' AND j_num = '" + j_num + "'");
        }

        public Boolean del_jrnl_entry(String j_code, String j_num)
        {
            return this.DeleteOnTable("tr02", "j_code = '" + j_code + "' AND j_num = '" + j_num + "'");
        }

        public DataTable ledger_supplier(String sl_code, DateTime asofdate, String invoice)
        {
            //return this.QueryBySQLCode("SELECT tr01.t_date, tr02.invoice, tr01.j_code, tr01.j_num, tr01.t_desc, tr02.debit, tr02.credit, tr02.seq_desc, tr01.fy, tr01.mo FROM " + schema + ".tr02 LEFT JOIN " + schema + ".tr01 ON (tr01.j_code=tr02.j_code AND tr01.j_num=tr02.j_num) WHERE tr01.cancel IS NULL AND tr02.sl_code = '000007' AND tr02.j_code != 'SJ' AND tr02.j_code != 'CRJ'  AND tr01.t_date <= '2014-05-24' AND (tr01.t_desc like '%00012929%' OR tr02.invoice like '%00012929%') ORDER BY tr01.t_date");
            return this.QueryBySQLCode("SELECT DISTINCT tr01.t_date, tr02.invoice, tr01.j_code, tr01.j_num, tr01.t_desc, SUM(tr02.debit), SUM(tr02.credit), tr02.seq_desc, tr01.fy, tr01.mo "
                                            + "FROM " + schema + ".tr02 LEFT JOIN " + schema + ".tr01 ON (tr01.j_code=tr02.j_code AND tr01.j_num=tr02.j_num) "
                                            + "WHERE tr01.cancel IS NULL AND tr02.j_code != 'SJ' AND tr02.j_code != 'CRJ' "
                                            + "AND tr02.sl_code='" + sl_code + "' AND tr01.t_date <= '" + asofdate.ToString("yyyy-MM-dd") + "' AND "
                                            + "tr02.invoice IN (SELECT DISTINCT tr02.invoice "
                                            + "FROM " + schema + ".tr02 LEFT JOIN " + schema + ".tr01 ON (tr01.j_code=tr02.j_code AND tr01.j_num=tr02.j_num) "
                                            + "WHERE tr01.cancel IS NULL AND tr02.j_code != 'SJ' AND tr02.j_code != 'CRJ' "
                                            + "AND tr02.sl_code='" + sl_code + "' AND tr01.t_date <= '" + asofdate.ToString("yyyy-MM-dd") + "' AND tr01.t_desc like '%" + invoice + "%' ) "
                                            + "GROUP BY tr01.t_date, tr02.invoice, tr01.j_code, tr01.j_num, tr01.t_desc, tr02.seq_desc, tr01.fy, tr01.mo "
                                            + "ORDER BY tr01.t_date ASC");
        }

        public DataTable ledger_customer(String sl_code, DateTime asofdate, String invoice)
        {
            //return this.QueryBySQLCode("SELECT tr01.t_date, tr02.invoice, tr01.j_code, tr01.j_num, tr01.t_desc, tr02.debit, tr02.credit, tr02.seq_desc, tr01.fy, tr01.mo FROM " + schema + ".tr02 LEFT JOIN " + schema + ".tr01 ON (tr01.j_code=tr02.j_code AND tr01.j_num=tr02.j_num) WHERE tr01.cancel IS NULL AND tr02.sl_code = '" + sl_code + "' AND tr02.j_code != 'PJ' AND tr02.j_code != 'CDJ'  AND tr01.t_date <= '" + asofdate.ToString("yyyy-MM-dd") + "'");
            return this.QueryBySQLCode("SELECT DISTINCT tr01.t_date, tr02.invoice, tr01.j_code, tr01.j_num, tr01.t_desc, SUM(tr02.debit), SUM(tr02.credit), tr02.seq_desc, tr01.fy, tr01.mo "
                                            + "FROM " + schema + ".tr02 LEFT JOIN " + schema + ".tr01 ON (tr01.j_code=tr02.j_code AND tr01.j_num=tr02.j_num) "
                                            + "WHERE tr01.cancel IS NULL AND tr02.j_code != 'PJ' AND tr02.j_code != 'CDJ' "
                                            + "AND tr02.sl_code='" + sl_code + "' AND tr01.t_date <= '" + asofdate.ToString("yyyy-MM-dd") + "' AND "
                                            + "tr02.invoice IN (SELECT DISTINCT tr02.invoice "
                                            + "FROM " + schema + ".tr02 LEFT JOIN " + schema + ".tr01 ON (tr01.j_code=tr02.j_code AND tr01.j_num=tr02.j_num) "
                                            + "WHERE tr01.cancel IS NULL AND tr02.j_code != 'PJ' AND tr02.j_code != 'CDJ' "
                                            + "AND tr02.sl_code='" + sl_code + "' AND tr01.t_date <= '" + asofdate.ToString("yyyy-MM-dd") + "' AND tr01.t_desc like '%" + invoice + "%' ) "
                                            + "GROUP BY tr01.t_date, tr02.invoice, tr01.j_code, tr01.j_num, tr01.t_desc, tr02.seq_desc, tr01.fy, tr01.mo "
                                            + "ORDER BY tr01.t_date ASC");
        }

        //for local site encoding rec_num starts 70000000, but from other sites rec_num starts 72000000
        public Boolean import_stktra(DataTable dt)
        {
            Boolean flag = false;
            String rec_num = "", new_rec_num = "";
            String rec_num_db = ""; //id to insert into db
            String item_code = "";
            Double qty = 0.00;
            String col_expiry = "", col_discount = "", col_cht_code = "", col_cnt_code = "", col_ln_vat = "", col_po_line = "";
            String val_expiry = "", val_discount = "", val_cht_code = "", val_cnt_code = "", val_ln_vat = "", val_po_line = "";

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                new_rec_num = dt.Rows[r]["rec_num"].ToString();

                if (String.IsNullOrEmpty(new_rec_num) == false && String.Compare(new_rec_num, rec_num) != 0)
                {
                    rec_num = new_rec_num;

                    rec_num_db = "I" + new_rec_num;
                    //insert rechdr
                    InsertOnTable("rechdr", "rec_num, reference, trnx_date, whs_code, recipient, t_date, t_time, trn_type, jrnlz, cancel", "'" + rec_num_db + "', $$" + dt.Rows[r]["reference"] + "$$, '" + dt.Rows[r]["trnx_date"] + "', '" + dt.Rows[r]["whs_code"] + "', '" + dt.Rows[r]["recipient"] + "', '" + dt.Rows[r]["t_date"] + "', '" + dt.Rows[r]["t_time"] + "', 'T', 'N', 'N'");
                }

                if (String.IsNullOrEmpty(new_rec_num) == false && String.Compare(new_rec_num, rec_num) == 0)
                {
                    item_code = dt.Rows[r]["item_code"].ToString();

                    if (String.IsNullOrEmpty(dt.Rows[r]["recv_qty"].ToString()) == false)
                    {
                        qty = Convert.ToDouble(dt.Rows[r]["recv_qty"].ToString());
                    }

                    adjust_item_qty_onhand(item_code, qty);

                    if (String.IsNullOrEmpty(dt.Rows[r]["discount"].ToString().Trim()) == false)
                    {
                        col_expiry = ", discount";
                        val_expiry = ", " + dt.Rows[r]["discount"];
                    }

                    if (String.IsNullOrEmpty(dt.Rows[r]["cht_code"].ToString().Trim()) == false)
                    {
                        col_expiry = ", cht_code";
                        val_expiry = ", '" + dt.Rows[r]["cht_code"] + "'";
                    }

                    if (String.IsNullOrEmpty(dt.Rows[r]["cnt_code"].ToString().Trim()) == false)
                    {
                        col_expiry = ", cnt_code";
                        val_expiry = ", '" + dt.Rows[r]["cnt_code"] + "'";
                    }

                    if (String.IsNullOrEmpty(dt.Rows[r]["ln_vat"].ToString().Trim()) == false)
                    {
                        col_expiry = ", ln_vat";
                        val_expiry = ", " + dt.Rows[r]["ln_vat"];
                    }

                    if (String.IsNullOrEmpty(dt.Rows[r]["po_line"].ToString().Trim()) == false)
                    {
                        col_expiry = ", po_line";
                        val_expiry = ", '" + dt.Rows[r]["po_line"] + "'";
                    }

                    if (String.IsNullOrEmpty(dt.Rows[r]["expiry"].ToString().Trim()) == false)
                    {
                        col_expiry = ", expiry";
                        val_expiry = ", '" + dt.Rows[r]["expiry"] + "'";
                    }

                    flag = InsertOnTable("reclne", "rec_num, ln_num, item_code, item_desc, unit, recv_qty, price, ln_amnt, lot_no" + col_discount + col_cht_code + col_cnt_code + col_ln_vat + col_po_line + col_expiry,
                        "'" + rec_num_db + "', '" + dt.Rows[r]["ln_num"] + "', '" + item_code + "', $$" + dt.Rows[r]["item_desc"] + "$$, '" + dt.Rows[r]["unit"] + "', " + qty.ToString("0.00") + ", " + dt.Rows[r]["price"] + ", " + dt.Rows[r]["ln_amnt"] + ", '" + dt.Rows[r]["lot_no"] + "'" + val_discount + val_cht_code + val_cnt_code + val_ln_vat + val_po_line + val_expiry);
                }
            }

            return flag;
        }

        public String get_next_stktra_rec_num_for_other_site()
        {
            String rec_num = "";

            try
            {
                DataTable dt = this.QueryOnTableWithParams("rechdr", "MAX(" + this.castToInteger("rec_num") + ")", "", "");

                rec_num = dt.Rows[0][0].ToString();

                rec_num = get_nextincrementlimitchar(rec_num, 8);
            }
            catch (Exception) { }

            return rec_num;
        }

        public Boolean insert_stkcrd()
        {
            Boolean flag = false;

            return flag;
        }

        public Boolean adjust_item_qty_onhand(String item_code, Double qty)
        {
            return UpdateOnTable("items", "qty_onhand=qty_onhand+" + qty + "", "item_code='" + item_code + "'");
        }

        //hotel
        public DataTable get_hotelguestfolioWithBalances(String gfolio, String gtype, Boolean includeWarehouse)
        {
            String WHERE = " AND gf.rom_code!='Z01' ";

            if (String.IsNullOrEmpty(gfolio) == false)
            {
                WHERE += " AND (gf.reg_num like '%" + gfolio + "%' OR gf.full_name like '%" + gfolio + "%')";
            }
            if (String.IsNullOrEmpty(gtype) == false)
            {
                WHERE += " AND rmrttyp='" + gtype + "'";
            }
            if (includeWarehouse)
            {
                WHERE += " AND gf.rom_code LIKE 'W%'";
            }

            return this.QueryBySQLCode("SELECT gf.reg_num, gf.full_name, gf.rom_code, gf.arr_date, gf.dep_date, gf.acct_no, gf.rmrttyp, SUM(ch.amount) AS balance FROM " + schema + ".gfolio gf LEFT JOIN " + schema + ".chgfil ch ON gf.reg_num=ch.reg_num LEFT JOIN " + schema + ".charge c ON c.chg_code=ch.chg_code WHERE c.isdeposit='FALSE'" + WHERE + " GROUP BY gf.reg_num, gf.full_name, gf.rom_code, gf.arr_date, gf.dep_date, gf.acct_no, gf.rmrttyp");
        }

        public DataTable get_guest_curchkin_selected(String reg_num)
        {
            return QueryBySQLCode("SELECT gf.rom_code, gf.typ_code, gf.full_name, gf.rmrttyp, gf.occ_type, gf.rate_code, gf.rom_rate, gf.govt_tax, gf.serv_chg, gf.arr_date, gf.dep_date, gf.reg_num, gf.user_id, gf.t_date, gf.t_time, gf.remarks, gf.bill_info, gf.rm_features, gf.disc_code, gf.disc_pct, g.address1 AS address, c.comp_name AS company FROM " + schema + ".gfolio gf LEFT JOIN " + schema + ".guest g ON g.acct_no=gf.acct_no LEFT JOIN " + schema + ".company c ON c.comp_code=g.comp_code WHERE cancel IS NULL AND reg_num='" + reg_num + "'");
            //return this.QueryOnTableWithParams("gfolio", "rom_code, typ_code, full_name, occ_type, rate_code, rom_rate, govt_tax, serv_chg, arr_date, dep_date, reg_num, user_id, t_date, t_time, remarks, bill_info, rm_features, disc_code, disc_pct", "cancel IS NULL AND reg_num='"+reg_num+"'", "");
        }

        public Double get_roomrateamt(String rttyp_cod, String rmtyp, int occtyp)
        {
            DataTable dt = new DataTable();
            String col_occ = "single"; //default

            try
            {
                if (occtyp == 2)
                {
                    col_occ = "double";
                }
                else if (occtyp == 3)
                {
                    col_occ = "triple";
                }
                else if (occtyp == 4)
                {
                    col_occ = "quad";
                }

                dt = this.QueryOnTableWithParams("romrate", col_occ, "rate_code='" + rttyp_cod + "' AND typ_code='" + rmtyp + "'", "");

                if (dt.Rows.Count > 0)
                {
                    return Convert.ToDouble(dt.Rows[0][0]);
                }
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }

            return 0.00;
        }

        public DataTable get_soalist(String search, String searchtype)
        {
            String WHERE = "";

            if (String.IsNullOrEmpty(search) == false)
            {
                if (searchtype == "soa_code")
                {
                    WHERE = "soa_code LIKE '%" + search + "%'";
                }
                else if (searchtype == "debt_name")
                {
                    WHERE = "debt_name LIKE '%" + search + "%'";
                }
                else
                {
                    WHERE = "soa_code LIKE '%" + search + "%' OR debt_name LIKE '%" + search + "%'";
                }
            }

            return this.QueryOnTableWithParams("soahdr2", "*", WHERE, " ORDER BY soa_code ASC");
        }

        public DataTable get_soalistWithBalance(String soa_code, String billedclient, int viewtype)
        {
            String WHERE = "";

            if (String.IsNullOrEmpty(soa_code) == false)
            {
                WHERE += " AND soa_code LIKE '%" + soa_code + "%'";
            }
            if (String.IsNullOrEmpty(billedclient) == false)
            {
                WHERE += " AND debt_name LIKE '%" + billedclient + "%'";
            }



            if (viewtype == 1)
            {
                //forCollection
                return this.QueryBySQLCode("SELECT gf.reg_num, gf.acct_no AS debt_code, gf.full_name AS debt_name, gf.rom_code, '0.00' AS amount, '' AS soa_code, '' AS soa_date, '' AS user_id, '' AS t_date, '' AS t_time, '' AS chg_dtfrm, '' AS chg_dtto FROM " + schema + ".gfolio gf WHERE gf.rom_code!='Z01' " + WHERE + " ORDER BY gf.rom_code");
            }
            else
            {
                return this.QueryBySQLCode("SELECT s.soa_code, s.debt_code, s.debt_name, s.soa_date, s.user_id, s.t_date, t_time,  s.chg_dtfrm, s.chg_dtto, SUM(sl.amount) AS amount, "
                                        + " (Select DISTINCT CASE WHEN gf.reg_num=s.fol_code THEN gf.rom_code ELSE gh.rom_code END  FROM rssys.gfolio gf, rssys.gfhist gh WHERE gf.reg_num=s.fol_code or gh.reg_num=s.fol_code) "
                                        + " FROM rssys.soahdr2 s LEFT JOIN rssys.soalne2 sl ON s.soa_code=sl.soa_code  "
                                        + " WHERE (s.cancel IS NULL OR s.cancel != 'Y') AND  (s.collected IS NULL OR s.collected=null OR s.collected != 'Y') " + WHERE
                                        + " GROUP BY  s.soa_code, s.debt_code, s.debt_name, s.soa_date, s.user_id, s.t_date, t_time,  s.chg_dtfrm, s.chg_dtto ORDER BY soa_code ASC");
            }


            return null;
            /*
            if (String.IsNullOrEmpty(soadtfrm) == false && String.IsNullOrEmpty(soadtto) == false)
            {
                WHERE += " AND s.soadtto LIKE '%" + search + "%'";
            }
            else if(String.IsNullOrEmpty(searchtype) == false)
            {
                WHERE = " AND soa_code LIKE '%" + search + "%' OR debt_name LIKE '%" + search + "%'";
            }*/


        }

        public Boolean is_check_m10(String code)
        {
            Boolean flag = false;
            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".m10 WHERE mp_code='" + code + "' AND ischeck='true';");
            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception) { }

            return flag;
        }

        public DataTable get_soalne_folio_list(String soa_code)
        {
            return this.QueryOnTableWithParams("soalne2", "*", "soa_code='" + soa_code + "'", " ORDER BY ln_num ASC");
        }

        public DataTable get_soalne_folio_list_report(String soa_code)
        {
            return this.QueryBySQLCode("SELECT s.*, gf.rom_code FROM " + schema + ".soalne3 s LEFT JOIN " + schema + ".gfolio gf ON s.reg_num=gf.reg_num WHERE soa_code='" + soa_code + "' ORDER BY ln_num ASC");
            //return this.QueryOnTableWithParams("soalne2", "*", "soa_code='" + soa_code + "'", " ORDER BY ln_num ASC");
        }

        public DataTable get_soalne_total_summary(String soa_code)
        {
            return this.QueryBySQLCode("SELECT s.chg_code, s.chg_desc, SUM(s.amount) AS amount FROM " + schema + ".soalne3 s WHERE soa_code='" + soa_code + "' GROUP BY s.chg_code, s.chg_desc");
            //return this.QueryOnTableWithParams("soalne2", "*", "soa_code='" + soa_code + "'", " ORDER BY ln_num ASC");
        }

        public DataTable get_soalne_foliotransaction_list(String soa_code)
        {
            return this.QueryOnTableWithParams("soalne3", "*", "soa_code='" + soa_code + "'", " ORDER BY reg_num ASC, ln_num ASC");
        }

        public Double get_soa_amt(String soa_code)
        {
            Double amt = 0;
            GlobalMethod gm = new GlobalMethod();

            try
            {
                DataTable dt = this.QueryOnTableWithParams("soalne2", "SUM(amount)", "soa_code='" + soa_code + "'", "");

                amt = gm.toNormalDoubleFormat(dt.Rows[0][0].ToString());
            }
            catch (Exception)
            {
                return 0;
            }

            return amt;
        }

        public String get_soa_reg_num(String soa_code)
        {
            DataTable dt = this.QueryOnTableWithParams("soahdr2", "fol_code", "soa_code='" + soa_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { }

            return val;
        }

        public DataTable get_collist(String search, String searchtype)
        {
            String WHERE = "";

            if (String.IsNullOrEmpty(search) == false)
            {
                if (searchtype == "or_code")
                {
                    WHERE = "or_code LIKE '%" + search + "%'";
                }
                else if (searchtype == "debt_name")
                {
                    WHERE = "debt_name LIKE '%" + search + "%'";
                }
            }

            return this.QueryOnTableWithParams("colhdr", "*", WHERE, " ORDER BY or_code ASC");
        }

        public DataTable get_collne_list(String or_code)
        {
            return this.QueryOnTableWithParams("collne", "*", "or_code='" + or_code + "'", " ORDER BY or_code ASC, ln_num ASC");
        }

        public DataTable get_collne2_list(String or_code)
        {
            return this.QueryOnTableWithParams("collne2", "*", "or_code='" + or_code + "'", " ORDER BY or_code ASC, ln_num ASC");
        }

        public DataTable get_guest_chargefil(String reg_num, Boolean isExceptDeposit)
        {
            String WHERE = "";

            if (isExceptDeposit)
            {
                WHERE = " AND c.isdeposit='FALSE'";
            }

            return this.QueryBySQLCode("SELECT cf.chg_date, cf.chg_code, c.chg_desc, cf.doc_type, CASE WHEN cf.chg_code = '102' OR cf.chg_code='103' OR cf.chg_code='104' OR cf.chg_code='105' THEN cf.reference ||' '|| cf.ccrd_no ||'-'|| cf.trace_no ELSE cf.reference END AS \"reference\", cf.amount AS \"amount\", cf.user_id, cf.t_time, cf.reference AS reference1, cf.ccrd_no, cf.trace_no, c.chg_type AS \"Type\", cf.chg_num  AS \"Chg Num\" FROM " + schema + ".chgfil cf INNER JOIN " + schema + ".charge c ON cf.chg_code=c.chg_code WHERE reg_num='" + reg_num + "'" + WHERE + " ORDER BY cf.t_date ASC, cf.t_time ASC, cf.chg_code ASC");
        }

        public DataTable get_guest_chargeOnly(String reg_num, String dtfrm, String dtto)
        {
            String WHERE = "";

            if (String.IsNullOrEmpty(dtfrm) == false && String.IsNullOrEmpty(dtto) == false)
            {
                WHERE = " AND cf.chg_date BETWEEN '" + dtfrm + "' AND '" + dtto + "'";
            }

            return this.QueryBySQLCode("SELECT cf.chg_date, cf.chg_code, c.chg_desc, cf.doc_type, CASE WHEN cf.chg_code = '102' OR cf.chg_code='103' OR cf.chg_code='104' OR cf.chg_code='105' THEN cf.reference ||' '|| cf.ccrd_no ||'-'|| cf.trace_no ELSE cf.reference END AS \"reference\", cf.amount AS \"amount\", cf.user_id, cf.t_time, cf.reference AS reference1, cf.ccrd_no, cf.trace_no, c.chg_type AS \"Type\", cf.chg_num  AS \"Chg Num\", cf.soa FROM " + schema + ".chgfil cf INNER JOIN " + schema + ".charge c ON cf.chg_code=c.chg_code WHERE reg_num='" + reg_num + "' AND c.chg_type='C'" + WHERE + " ORDER BY cf.t_date ASC, cf.t_time ASC, cf.chg_code ASC");
        }

        public Boolean is_chg_deposit(String chg_code)
        {
            DataTable dt = new DataTable();

            try
            {
                dt = this.QueryOnTableWithParams("charge", "1", "chg_code='" + chg_code + "' AND isdeposit='TRUE'", "");

                if (dt.Rows.Count > 0)
                {
                    return true;
                }
            }
            catch (Exception) { }

            return false;
        }

        public Double get_guest_charges_total(String reg_num)
        {
            Double amt = 0.00;
            DataTable dt = new DataTable();
            String WHERE = "";

            try
            {
                dt = this.QueryBySQLCode("SELECT SUM(cf.amount) FROM " + schema + ".chgfil cf INNER JOIN " + schema + ".charge c ON cf.chg_code=c.chg_code WHERE reg_num='" + reg_num + "' AND c.chg_type='C'" + WHERE);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        amt = Convert.ToDouble(row[0].ToString());
                    }
                }
            }
            catch (Exception) { }

            return amt;
        }

        public Double get_guest_payments_total(String reg_num, Boolean isExceptDeposit)
        {
            Double amt = 0.00;
            DataTable dt = new DataTable();
            String WHERE = "";

            try
            {
                if (isExceptDeposit)
                {
                    WHERE = " AND c.isdeposit='FALSE'";
                }

                dt = this.QueryBySQLCode("SELECT SUM(cf.amount) FROM " + schema + ".chgfil cf INNER JOIN " + schema + ".charge c ON cf.chg_code=c.chg_code WHERE reg_num='" + reg_num + "' AND c.chg_type='P'" + WHERE);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        amt = Convert.ToDouble(row[0].ToString());
                    }
                }
            }
            catch (Exception) { }

            return amt;
        }

        public Double get_guest_charge_deposit(String reg_num)
        {
            Double amt = 0.00;
            DataTable dt = new DataTable();

            try
            {

                dt = this.QueryBySQLCode("SELECT SUM(cf.amount) FROM " + schema + ".chgfil cf INNER JOIN " + schema + ".charge c ON cf.chg_code=c.chg_code WHERE reg_num='" + reg_num + "' AND c.isdeposit='TRUE'");
                //("chgfil", "SUM(amount)", "reg_num='" + reg_num + "'", "");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        amt = Convert.ToDouble(row[0].ToString());
                    }
                }
            }
            catch (Exception) { }

            return amt;
        }

        //inventory
        public Double get_itemregsellprice(String itemcode)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "sell_pric", "item_code='" + itemcode + "'", "");
            Double price = 0.00;

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    price = gm.toNormalDoubleFormat(dt.Rows[i][0].ToString());
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return price;
        }

        public String get_itemsellunitid(String itemcode)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "sales_unit_id", "item_code='" + itemcode + "'", "");
            String si = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    si = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return si;
        }

        public DataTable get_itemsearchlist(String srch)
        {
            return this.QueryOnTableWithParams("items", "qty_onhand_su, item_desc, sales_unit_id AS sale_unit, sell_pric as regular, sc_price AS senior, bin_loc, item_code", "item_desc LIKE '%" + srch + "%' OR item_code='%" + srch + "%'", "ORDER BY item_desc ASC");
        }

        //type: Q =
        public DataTable get_itemlist(String search, String typ, String sort1, String sort2)
        {
            String WHERE = "";
            String SORT = "ig.grp_desc, i.item_desc ";

            if (String.IsNullOrEmpty(search) == false && typ == "C")
            {
                WHERE = " WHERE i.item_code LIKE '%" + search + "%'";
            }
            else if (String.IsNullOrEmpty(search) == false && typ == "Q")
            {
                WHERE = " WHERE i.qty_onhand_su LIKE '%" + search + "%'";
            }
            else if (String.IsNullOrEmpty(search) == false && typ == "B")
            {
                WHERE = " WHERE b.brd_name LIKE '%" + search + "%'";
            }
            else if (String.IsNullOrEmpty(search) == false && typ == "G")
            {
                WHERE = " WHERE ig.grp_desc LIKE '%" + search + "%'";
            }
            else if (String.IsNullOrEmpty(search) == false)
            {
                WHERE = " WHERE i.item_desc LIKE '%" + search + "%'";
            }
            /*
            /// sort 1
            if (String.IsNullOrEmpty(search) == false && sort1 == "C")
            {
                SORT = " i.item_code ";
            }
            else if (String.IsNullOrEmpty(search) == false && sort1 == "Q")
            {
                SORT = " i.qty_onhand_su";
            }
            else if (String.IsNullOrEmpty(search) == false && sort1 == "B")
            {
                SORT = " b.brd_name";
            }
            else if (String.IsNullOrEmpty(search) == false && sort1 == "G")
            {
                SORT = " ig.grp_desc";
            }
            else if (String.IsNullOrEmpty(search) == false)
            {
                SORT = " i.item_desc";
            }

            /// sort 2
            if (String.IsNullOrEmpty(search) == false && sort2 == "C")
            {
                SORT += " i.item_code ";
            }
            else if (String.IsNullOrEmpty(search) == false && sort2 == "Q")
            {
                SORT += " i.qty_onhand_su";
            }
            else if (String.IsNullOrEmpty(search) == false && sort2 == "B")
            {
                SORT += " b.brd_name";
            }
            else if (String.IsNullOrEmpty(search) == false && sort2 == "G")
            {
                SORT += " ig.grp_desc";
            }
            else if (String.IsNullOrEmpty(search) == false)
            {
                SORT += " i.item_desc";
            }
            */

            return this.QueryBySQLCode("SELECT i.qty_onhand_su, i.item_desc, iu.unit_shortcode AS sale_unit, i.sales_unit_id, b.brd_name, i.brd_code, i.sell_pric AS regular, i.sc_price AS senior, i.bin_loc, i.item_code, i.unit_cost, ig.grp_desc FROM " + schema + ".items i LEFT JOIN " + schema + ".itmunit iu ON i.sales_unit_id=iu.unit_id LEFT JOIN " + schema + ".brand b ON i.brd_code=b.brd_code LEFT JOIN " + schema + ".itmgrp ig ON ig.item_grp=i.item_grp " + WHERE + " ORDER BY "+SORT+"  ASC");
        }

        public DataTable get_solist(DateTime frm, DateTime to, Boolean pendingOnly)
        {
            return this.QueryBySQLCode("SELECT o.*, w.* FROM " + schema + ".orhdr o LEFT JOIN  " + schema + ".whouse w ON w.whs_code=o.loc WHERE trnx_date BETWEEN '" + frm.ToString("yyyy-MM-dd") + "' AND '" + to.ToString("yyyy-MM-dd") + "'");
        }

        public DataTable get_so_code_ByItem(String itemcode)
        {
            return this.QueryBySQLCode("SELECT DISTINCT ol.*, o.inv_num, o.rep_code FROM " + schema + ".orhdr o INNER JOIN " + schema + ".orlne ol ON o.ord_code=ol.ord_code WHERE ol.item_code='" + itemcode + "' ORDER BY ol.ord_code ASC");
        }

        public DataTable get_soretlist(DateTime frm, DateTime to)
        {
            return this.QueryBySQLCode("SELECT o.*, w.* FROM " + schema + ".rethdr o LEFT JOIN  " + schema + ".whouse w ON w.whs_code=o.whs_code WHERE o.ret_date BETWEEN '" + frm.ToString("yyyy-MM-dd") + "' AND '" + to.ToString("yyyy-MM-dd") + "'");
        }

        public DataTable get_soret_itemlist(String ret_num)
        {
            return this.QueryBySQLCode("SELECT r.*, u.* FROM " + schema + ".retlne r LEFT JOIN " + schema + ".itmunit u ON r.unit=u.unit_id WHERE r.ret_num='" + ret_num + "' ");
        }
        
        public DataTable get_itemForSaleEntry(String itemcode)
        {
            return this.QueryOnTableWithParams("items", "*", "item_code='" + itemcode + "'", "");
        }

        public DataTable get_polist(DateTime frm, DateTime to, String typ)
        {
            var query = "SELECT p.inv_num, p.supl_code, p.supl_name, p.reference, TO_CHAR(p.inv_date, 'MM/DD/YYYY') AS inv_date, p.whs_code, p.pay_code, p.trn_type, p.jrnlz, p.cancel, p.userid, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time, p.final, w.whs_desc, mop.mp_desc FROM " + schema + ".pinvhd p LEFT JOIN " + schema + ".m07 s ON s.c_code=p.supl_code LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code LEFT JOIN " + schema + ".m10 mop ON mop.mp_code=p.pay_code WHERE  p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' AND p.trn_type='" + typ.ToUpper() + "' ORDER BY inv_num";
            // MessageBox.Show("SELECT p.inv_num, p.supl_code, p.supl_name, p.reference, TO_CHAR(p.inv_date, 'MM/DD/YYYY') AS inv_date, p.whs_code, p.pay_code, p.trn_type, p.jrnlz, p.cancel, p.userid, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time, p.final, w.whs_desc, mop.mp_desc FROM " + schema + ".pinvhd p LEFT JOIN " + schema + ".m07 s ON s.c_code=p.supl_code LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code LEFT JOIN " + schema + ".m10 mop ON mop.mp_code=p.pay_code WHERE  p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' AND p.trn_type='" + typ.ToUpper() + "' ORDER BY inv_num");
            return this.QueryBySQLCode("SELECT p.prec_num, p.inv_num, p.supl_code, p.supl_name, p.reference, TO_CHAR(p.inv_date, 'MM/DD/YYYY') AS inv_date, p.whs_code, p.pay_code, p.trn_type, p.jrnlz, p.cancel, p.userid, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time, p.final, w.whs_desc, mop.mp_desc FROM " + schema + ".pinvhd p LEFT JOIN " + schema + ".m07 s ON s.c_code=p.supl_code LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code LEFT JOIN " + schema + ".m10 mop ON mop.mp_code=p.pay_code WHERE  p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' AND p.trn_type='" + typ.ToUpper() + "' ORDER BY inv_num");
        }

        public DataTable get_po_item_list(String inv_num)
        {
            return this.QueryOnTableWithParams("pinvln", "*", "inv_num='" + inv_num + "'", "ORDER BY " + castToInteger("ln_num") + "");
            //return this.QueryOnTableWithParams("reclne", "*", "rec_num='" + inv_num + "'", "ORDER BY " + castToInteger("ln_num") + "");
        }

        public DataTable get_retpolist(DateTime frm, DateTime to)
        {
            return this.QueryBySQLCode("SELECT p.pret_num, p.supl_code, p.supl_name, p.reference, TO_CHAR(p.trnx_date, 'MM/DD/YYYY') AS trnx_date, p.whs_code, p.jrnlz, p.cancel, p.user_id, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time FROM " + schema + ".prethdr p LEFT JOIN " + schema + ".m07 s ON s.c_code=p.supl_code LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code WHERE  p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' ORDER BY pret_num");
        }

        public DataTable get_retpo_item_list(String pret_num)
        {
            return this.QueryOnTableWithParams("pretlne", "*", "pret_num='" + pret_num + "'", "ORDER BY " + castToInteger("ln_num") + "");
        }

        public DataTable get_stkinvlist(DateTime frm, DateTime to, String typ) //for stocks alone
        {
            var item = "SELECT p.\"locationFrom\", p.\"locationTo\", p.rec_num, p.supl_name, p._reference, TO_CHAR(p.trnx_date, 'MM/DD/YYYY') AS trnx_date, p.whs_code, p.trn_type, p.jrnlz, p.cancel, p.recipient, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time, p.vat_code, w.whs_desc, p.purc_ord, p.printed FROM " + schema + ".rechdr p LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code WHERE  p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' AND p.trn_type='" + typ.ToUpper() + "' ORDER BY rec_num";

            return this.QueryBySQLCode("SELECT p.\"locationFrom\", p.\"locationTo\", p.rec_num, p.supl_name,p.supl_code, p._reference, TO_CHAR(p.trnx_date, 'MM/DD/YYYY') AS trnx_date, p.whs_code, p.trn_type, p.jrnlz, p.cancel, p.recipient, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time, p.vat_code, w.whs_desc, p.purc_ord, p.printed FROM " + schema + ".rechdr p LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code WHERE  p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' AND p.trn_type='" + typ.ToUpper() + "' ORDER BY rec_num");
            //            return this.QueryBySQLCode("SELECT p.rec_num, p.supl_code, p.supl_name, p.reference, TO_CHAR(p.trnx_date, 'MM/DD/YYYY') AS trnx_date, p.whs_code, p.trn_type, p.jrnlz, p.cancel, p.recipient, TO_CHAR(p.t_date, 'MM/DD/YYYY') AS t_date, p.t_time, p.vat_code, w.whs_desc, p.purc_ord, p.printed FROM " + schema + ".rechdr p LEFT JOIN " + schema + ".m07 s ON s.c_code=p.supl_code LEFT JOIN " + schema + ".whouse w ON w.whs_code=p.whs_code WHERE (p.cancel IS NULL OR p.cancel='N') AND p.t_date >= '" + frm.ToString("yyyy-MM-dd") + "' AND p.t_date <= '" + to.ToString("yyyy-MM-dd") + "' AND p.trn_type='" + typ.ToUpper() + "' ORDER BY rec_num");
        }

        public DataTable get_stkinv_item_list(String rec_num) //for stocks alone
        {
            return this.QueryBySQLCode("SELECT r.*, i.unit_desc FROM " + this.schema + ".reclne r LEFT JOIN " + this.schema + ".itmunit i ON r.unit=i.unit_id WHERE rec_num='" + rec_num + "' ORDER BY " + castToInteger("ln_num") + "");
            //            return this.QueryBySQLCode("SELECT r.*, i.unit_desc FROM " + this.schema + ".reclne r LEFT JOIN " + this.schema + ".itmunit i ON r.unit_id=i.unit_id WHERE rec_num='" + rec_num + "' ORDER BY " + castToInteger("ln_num") + "");
        }

        public DataTable get_stkinv_item_by_item_code(String itemCode) //for stocks alone
        {
            return this.QueryBySQLCode("SELECT s.item_code, s.item_desc, s.unit, TO_CHAR(s.trnx_date, 'MM/DD/YYYY') AS trnx_date, s.reference, s.po_so, s.qty_in, s.qty_out, s.fcp, price, s.whs_code, s.supl_code, s.supl_name, s.trn_type, s.cht_code, s.cnt_code, i.unit_shortcode, w.whs_desc FROM " + this.schema + ".stkcrd s LEFT JOIN " + this.schema + ".itmunit i ON s.unit=i.unit_id LEFT JOIN " + this.schema + ".whouse w ON w.whs_code=s.whs_code WHERE item_code='" + itemCode + "' ORDER BY s.trnx_date, trn_type;");
        }

        public Boolean is_itemcode(String code)
        {
            Boolean flag = false;
            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".items WHERE item_code='" + code + "';");
            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception) { }

            return flag;
        }

        public String get_customer_name(String code)
        {
            DataTable dt = this.QueryOnTableWithParams("m06", "d_name", "d_code='" + code + "'", "");
            String cname = null;

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cname = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return cname;
        }
        /*
        public Double get_itemregsellprice(String itemcode)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "sell_pric", "item_code='" + itemcode + "'", "");
            Double price = 0.00;

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    price = gm.toNormalDoubleFormat(dt.Rows[i][0].ToString());
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return price;
        }

        public String get_itemsellunitid(String itemcode)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "sales_unit_id", "item_code='" + itemcode + "'", "");
            String si = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    si = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return si;
        }*/

        public DataTable get_collector_list()
        {
            return this.QueryOnTableWithParams("collector", "*", "", "ORDER BY coll_code ASC");
        }

        public DataTable get_customer_list()
        {
            return this.QueryOnTableWithParams("m06", "*", "", "ORDER BY d_code ASC");
        }

        public DataTable get_salesrep_list()
        {
            return this.QueryOnTableWithParams("repmst", "*", "", "ORDER BY rep_code ASC");
        }

        public DataTable get_salesrep_list_with_canc()
        {
            return this.QueryOnTableWithParams("repmst", "*", "", "ORDER BY rep_code ASC");
        }

        public String get_salesrep_name(String rep_code)
        {
            DataTable dt = this.QueryOnTableWithParams("repmst", "rep_name", "rep_code='" + rep_code + "'", "ORDER BY rep_code ASC");
            String name = "";

            if (dt.Rows.Count > 0)
            {
                name = dt.Rows[0]["rep_name"].ToString();
            }

            return name;
        }

        public Boolean is_salesrep_password_duplicate(String pass)
        {
            Boolean flag = false;

            DataTable dt = this.QueryBySQLCode("SELECT 1 FROM " + schema + ".repmst WHERE pwd='" + pass + "';");

            try
            {
                if (dt.Rows[0][0].ToString() == "1")
                {
                    flag = true;
                }
            }
            catch (Exception) { }

            return flag;
        }

        public DataTable get_generic_list()
        {
            return this.QueryOnTableWithParams("generic", "*", "cancel IS NULL OR cancel='N'", "ORDER BY gen_code ASC");
        }

        public DataTable get_rType_list()
        {
            return this.QueryOnTableWithParams("rtype", "*", "", "ORDER BY typ_code ASC");
        }

        public DataTable get_brand_list()
        {
            return this.QueryOnTableWithParams("brand", "*", "cancel IS NULL OR cancel='N'", "ORDER BY brd_code ASC");
        }

        public DataTable get_item_details(String code)
        {
            return this.QueryOnTableWithParams("items", "*", "item_code='" + code + "' ", "");
        }

        public DataTable get_item_details_multiple(String code, int noOfItems)
        {
            String WHERE = "";
            String sql = "";

            if(String.IsNullOrEmpty(code) == false)
            {
                WHERE = " WHERE item_code='" + code + "' ";
            }

            sql = "SELECT * FROM " + this.schema + ".items" + WHERE;

            for (int i = 1; i < noOfItems; i++ )
            {
                sql = sql + " UNION ALL SELECT * FROM " + this.schema + ".items " + WHERE;
            }

            sql = sql + " ORDER BY item_code ASC;";
            
            return this.QueryBySQLCode(sql);
        }

        public Boolean save_fcp_to_items(String item_code, Double fcp)
        {
            return this.UpdateOnTable("items", "fcp='" + fcp.ToString("0.00") + "'", "item_code='" + item_code + "'");
        }

        public Boolean save_ave_cost_to_items(String item_code, Double ave_cost)
        {
            return this.UpdateOnTable("items", "ave_cost='" + ave_cost.ToString("0.00") + "'", "item_code='" + item_code + "'");
        }

        public Boolean save_to_stkcard(String item_code, String item_desc, String unit, String trnx_date, String reference, String po_so, String qty_in, String qty_out, String fcp, String price, String whs_code, String supl_code, String supl_name, String trn_type, String cht_code, String cnt_code, String proj_code)
        {
            Boolean flag = false, flag2 = false;
            String stk_qty_in = "0.00";
            String stk_qty_out = "0.00";

            fcp = get_item_fcp_compute(item_code).ToString("0.00");

            flag2 = save_fcp_to_items(item_code, gm.toNormalDoubleFormat(fcp));
            flag2 = save_ave_cost_to_items(item_code, get_item_ave_cost_compute(item_code));

            //qty_in
            if (gm.toNormalDoubleFormat(qty_in) < 0)
                stk_qty_out = Math.Abs(gm.toNormalDoubleFormat(qty_in)).ToString();
            else
                stk_qty_in = qty_in;

            //qty_out
            if (gm.toNormalDoubleFormat(qty_out) < 0)
                stk_qty_in = Math.Abs(gm.toNormalDoubleFormat(qty_out)).ToString();
            else
                stk_qty_out = qty_out;

            if (this.isExists("stkcrd", "po_so='" + po_so + "' AND trn_type='" + trn_type + "'") == false)
            {
                flag = this.InsertOnTable("stkcrd", "item_code, item_desc, unit, trnx_date, reference, po_so, qty_in, qty_out, fcp, price, whs_code, supl_code, supl_name, trn_type, cht_code, cnt_code,  proj_code", "'" + item_code + "', " + this.str_E(item_desc) + ", '" + unit + "', '" + trnx_date + "', " + this.str_E(reference) + ", '" + po_so + "', '" + stk_qty_in + "', '" + stk_qty_out + "', '" + fcp + "', '" + price + "', '" + whs_code + "', '" + supl_code + "', '" + supl_name + "', '" + trn_type + "', '" + cht_code + "', '" + cnt_code + "', '" + proj_code + "'");
            }
            else
            {
                flag = this.UpdateOnTable("stkcrd", "item_code='" + item_code + "', item_desc='" + item_desc + "', unit='" + unit + "', trnx_date='" + trnx_date + "', reference='" + reference + "', po_so='" + po_so + "', qty_in='" + stk_qty_in + "', qty_out='" + stk_qty_out + "', fcp='" + fcp + "', price='" + price + "', whs_code='" + whs_code + "', supl_code='" + supl_code + "', supl_name='" + supl_name + "', trn_type='" + trn_type + "', cht_code='" + cht_code + "', cnt_code='" + cnt_code + "',  proj_code='" + proj_code + "'", "po_so='" + po_so + "' AND trn_type='" + trn_type + "'");
            }

            return flag;
        }
        
        //typ_io = P-purchase, S-sales , SR - Sales Return, PR - Purchase Return, T - transfer, I - Issuance, A - Adjustment
        public Boolean upd_item_quantity_onhand(String itemcode, Double qty, String typ_ps)
        {
            Boolean success = false;
            GlobalClass.DontSendToMain = true;

            if (typ_ps == "P" || typ_ps == "SR" || typ_ps == "A" || typ_ps == "T" || typ_ps == "E" || typ_ps == "R")
            {
                //addition
                success = this.UpdateOnTable("items", "qty_onhand_su = " + (this.get_item_qty(itemcode) + qty).ToString("0.00"), "item_code='" + itemcode + "'");
            }
            else
            {
                //subtraction
                success = this.UpdateOnTable("items", "qty_onhand_su = " + (this.get_item_qty(itemcode) - qty).ToString("0.00"), "item_code='" + itemcode + "'");
            }

            GlobalClass.DontSendToMain = false;

            return success;
        }
        public String get_item_desc(String code)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "item_desc", "item_code='" + code + "'", "");
            String item_desc = "";

            try
            {
                if (dt.Rows.Count > 0)
                {
                    item_desc = dt.Rows[0]["item_desc"].ToString();
                }
            }
            catch (Exception) { }

            return item_desc;
        }

        public Double get_item_cost(String code)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "unit_cost", "item_code='" + code + "'", "");
            Double cost = 0.00;

            try
            {
                if (dt.Rows.Count > 0)
                {
                    cost = Convert.ToDouble(dt.Rows[0]["unit_cost"].ToString());
                }
            }
            catch (Exception) { }

            return cost;
        }

        public String get_item_costcode(String code)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "ccode", "item_code='" + code + "'", "");
            String ccode = "";

            try
            {
                if (dt.Rows.Count > 0)
                {
                    ccode = dt.Rows[0]["ccode"].ToString();
                }
            }
            catch (Exception) { }

            return ccode;
        }

        public Double get_item_qty(String itemcode)
        {
            DataTable dt = this.QueryOnTableWithParams("items", "qty_onhand_su", "item_code='" + itemcode + "'", "");
            Double qty_onhand_su = 0.00;

            try
            {
                if (dt.Rows.Count > 0)
                {
                    qty_onhand_su = Convert.ToDouble(dt.Rows[0]["qty_onhand_su"].ToString());
                }
            }
            catch (Exception) { }

            return qty_onhand_su;
        }

        public String get_item_unit_desc(String id)
        {
            DataTable dt = this.QueryOnTableWithParams("itmunit", "unit_desc", "unit_id='" + id + "'", "");
            String unit_desc = null;

            try
            {
                if (dt.Rows.Count > 0)
                {
                    unit_desc = dt.Rows[0]["unit_desc"].ToString();
                }
            }
            catch (Exception) { }

            return unit_desc;
        }

        public String get_costcenter_desc(String cc_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m08", "cc_desc", "cc_code='" + cc_code + "'", "");
            String cc_desc = null;

            try
            {
                if (dt.Rows.Count > 0)
                {
                    cc_desc = dt.Rows[0]["cc_desc"].ToString();
                }
            }
            catch (Exception) { }

            return cc_desc;
        }

        public DataTable get_vat_list()
        {
            return this.QueryOnTableWithParams("vat", "*", "", "ORDER BY vat_code ASC");
        }

        public String get_vat_desc(String vat_code)
        {
            DataTable dt = new DataTable();
            String desc = "";

            try
            {
                dt = this.QueryOnTableWithParams("vat", "vat_desc", "vat_code='" + vat_code + "'", "");

                desc = dt.Rows[0]["vat_desc"].ToString();
            }
            catch (Exception) { }

            return desc;
        }

        public Double get_vat_pct(String vat_code)
        {
            DataTable dt = new DataTable();
            Double pct = 0.00;

            try
            {
                dt = this.QueryOnTableWithParams("vat", "vat_pct", "vat_code='" + vat_code + "'", "");

                pct = gm.toNormalDoubleFormat(dt.Rows[0]["vat_pct"].ToString());
            }
            catch (Exception) { }

            return pct;
        }

        //return array of String [0] = clerk id, [1] = clerk name; otherwise null
        public String[] verifyClerk(String pass)
        {
            DataTable dt = this.QueryOnTableWithParams("repmst", "rep_code, rep_name", "pwd=$$" + pass + "$$", "");
            String[] clerk;

            clerk = new String[2];

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clerk[0] = dt.Rows[i][0].ToString();
                    clerk[1] = dt.Rows[i][1].ToString();
                }
            }
            catch (Exception) { }

            return clerk;
        }

        public String verifyUserIDForDiscount(String pass)
        {
            DataTable dt = this.QueryOnTableWithParams("x08", "uid", "pwd=$$" + pass + "$$ AND grp_id='001'", "");
            String clerk = null;

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clerk = dt.Rows[i]["uid"].ToString();
                }
            }
            catch (Exception) {  }

            return clerk;
        }

        public String get_m99fy()
        {
            DataTable dt = this.QueryOnTableWithParams("m99", "fy", "", "");
            String fy = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    fy = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return fy;
        }

        public String get_default_whs()
        {
            DataTable dt = this.QueryOnTableWithParams("m99", "def_whs", "", "");
            String def_whs = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    def_whs = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return def_whs;
        }

        public String get_j_desc(String j_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m05", "j_desc", "j_code='" + j_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { }

            return val;
        }

        public String get_checkby(String j_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m05", "check_by", "j_code='" + j_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { }

            return val;
        }

        public String get_approvby(String j_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m05", "approv_by", "j_code='" + j_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { }

            return val;
        }

        public String get_notedby(String j_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m05", "noted_by", "j_code='" + j_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { }

            return val;
        }

        public String get_preparedby(String j_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m05", "prt_type", "j_code='" + j_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception er) { }

            return val;
        }

        public Double get_paidtoamount(String j_code, String j_num)
        {
            GlobalMethod gm = new GlobalMethod();
            DataTable dt = this.QueryOnTableWithParams("tr02", "debit, credit, sl_code", "j_code='" + j_code + "' AND j_num='" + j_num + "'", "");
            Double val = 0.00, dr = 0.00, cr = 0.00;

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //if (String.IsNullOrEmpty(dt.Rows[i]["sl_code"].ToString()) == false)
                    //{
                    dr = gm.toNormalDoubleFormat(dt.Rows[i]["debit"].ToString());
                    cr = gm.toNormalDoubleFormat(dt.Rows[i]["credit"].ToString());
                    //}
                }
            }
            catch (Exception er) { }

            if (dr > 0)
            {
                val = dr;
            }
            else if (cr > 0)
            {
                val = cr;
            }
            return val;
        }

        //purchase request
        public DataTable get_purhace_request_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT pr.*, c.*, sc.*, (SELECT DISTINCT pl.date_need FROM " + schema + ".prlne pl WHERE pl.pr_code=pr.pr_code) AS date_need FROM " + schema + ".prhdr pr LEFT JOIN " + schema + ".m08 c ON pr.cnt_code=c.cc_code LEFT JOIN " + schema + ".subctr sc ON sc.scc_code=pr.scc_code WHERE pr.pr_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' ORDER BY pr.pr_date");
        }

        public DataTable get_purhace_request_byprcode(String pr_code_frm, String pr_code_to)
        {
            return this.QueryBySQLCode("SELECT pr.*, c.*, sc.* FROM " + schema + ".prhdr pr LEFT JOIN " + schema + ".m08 c ON pr.cnt_code=c.cc_code LEFT JOIN " + schema + ".subctr sc ON sc.scc_code=pr.scc_code WHERE pr.pr_code BETWEEN '" + pr_code_frm + "' AND '" + pr_code_to + "' ORDER BY pr.pr_code");
        }

        public DataTable get_purhace_request_byitem(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT pr.*, prl.* c.*, sc.* FROM " + schema + ".prhdr pr LEFT JOIN " + schema + ".m08 c ON pr.cnt_code=c.cc_code LEFT JOIN " + schema + ".subctr sc ON sc.scc_code=pr.scc_code  LEFT JOIN " + schema + ".prlne prl ON pr.pr_code=prl.pr_code WHERE pr.pr_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' ORDER BY prl.pr_code, prl.ln_num");
        }

        public DataTable get_purhace_request_byscc(String scc_code_frm, String scc_code_to)
        {
            return this.QueryBySQLCode("SELECT pr.*, c.*, sc.* FROM " + schema + ".prhdr pr LEFT JOIN " + schema + ".m08 c ON pr.cnt_code=c.cc_code LEFT JOIN " + schema + ".subctr sc ON sc.scc_code=pr.scc_code WHERE pr.scc_code BETWEEN '" + scc_code_frm + "' AND '" + scc_code_to + "' ORDER BY pr.scc_code, pr.pr_code");
        }
        public DataTable get_purhace_request_items(String pr_code)
        {
            return this.QueryBySQLCode("SELECT pl.*, c.*, sc.*, u.*, (pl.quantity * pl.price) AS ln_amnt FROM  " + schema + ".prhdr pr LEFT JOIN " + schema + ".prlne pl ON pr.pr_code=pl.pr_code LEFT JOIN " + schema + ".m08 c ON pr.cnt_code=c.cc_code LEFT JOIN " + schema + ".subctr sc ON sc.scc_code=pr.scc_code LEFT JOIN " + schema + ".itmunit u ON u.unit_id=pl.purc_unit WHERE pl.pr_code='" + pr_code + "' ORDER BY " + castToInteger("pl.ln_num") + "");
        }

        //purchase order
        public DataTable get_po_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT pr.*, m.*, m7.* FROM " + schema + ".purhdr pr LEFT JOIN " + schema + ".m10 m ON pr.pay_code=m.mp_code LEFT JOIN " + schema + ".m07 m7 ON pr.supl_name=m7.c_code WHERE pr.t_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' ORDER BY pr.t_date");
        }

        public DataTable get_po_bynumber(String nofrm, String noto)
        {
            return this.QueryBySQLCode("SELECT pr.*, m.*, m7.* FROM " + schema + ".purhdr pr LEFT JOIN " + schema + ".m10 m ON pr.pay_code=m.mp_code LEFT JOIN " + schema + ".m07 m7 ON pr.supl_name=m7.c_code WHERE pr.purc_ord BETWEEN '" + nofrm + "' AND '" + noto + "' ORDER BY pr.purc_ord");
        }

        public DataTable get_po_bysupplier(String supfrm, String supto)
        {
            return this.QueryBySQLCode("SELECT pr.*, m7.c_name, prl.* FROM " + schema + ".purhdr pr LEFT JOIN " + schema + ".m07 m7 ON pr.supl_name=m7.c_code LEFT JOIN " + schema + ".purlne prl ON pr.purc_ord=prl.purc_ord WHERE pr.supl_name BETWEEN '" + supfrm + "' AND '" + supto + "' ORDER BY pr.supl_name");
        }

        public DataTable get_po_byitems(String itmfrm, String itmto)
        {
            return this.QueryBySQLCode("SELECT pr.*, m7.c_name, prl.* FROM " + schema + ".purhdr pr LEFT JOIN " + schema + ".m07 m7 ON pr.supl_name=m7.c_code LEFT JOIN " + schema + ".purlne prl ON pr.purc_ord=prl.purc_ord WHERE prl.item_code BETWEEN '" + itmfrm + "' AND '" + itmto + "' ORDER BY prl.purc_ord, prl.item_code");
        }
        
        public DataTable get_po_items(String purc_ord)
        {
            return this.QueryBySQLCode("SELECT pr.*, pl.*, c.*, sc.*, u.* FROM  " + schema + ".purhdr pr LEFT JOIN " + schema + ".purlne pl ON pr.purc_ord=pl.purc_ord LEFT JOIN " + schema + ".m08 c ON pl.cnt_code=c.cc_code LEFT JOIN " + schema + ".subctr sc ON sc.scc_code=pl.scc_code LEFT JOIN " + schema + ".itmunit u ON u.unit_id=pl.purc_unit WHERE pl.purc_ord='" + purc_ord + "' ORDER BY " + castToInteger("pl.ln_num") + "");
        }

        //receiving
        public DataTable get_rr_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.trnx_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' ORDER BY r.trnx_date");
        }

        public DataTable get_rr_bynumber(String nofrm, String noto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.purc_ord BETWEEN '" + nofrm + "' AND '" + noto + "' ORDER BY r.rec_num");
        }

        public DataTable get_rr_byitems(String itmfrm, String itmto)
        {
            return this.QueryBySQLCode("SELECT r.*, rl.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE rl.item_code BETWEEN '" + itmfrm + "' AND '" + itmto + "' ORDER BY rl.rec_num, rl.item_code");
        }

        public DataTable get_rr_items(String rec_num)
        {
            //return this.QueryBySQLCode("SELECT rl.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE rl.rec_num='" + rec_num + "' ORDER BY rl.rec_num, rl.item_code");
            return this.QueryBySQLCode("SELECT rl.*, u.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num LEFT JOIN " + schema + ".itmunit u ON u.unit_id=rl.unit WHERE rl.rec_num='" + rec_num + "' ORDER BY rl.rec_num, rl.item_code");
        }

        //direct purchase 
        public DataTable get_dr_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT dr.*, w.*, m7.*, m.* FROM " + schema + ".pinvhd dr LEFT JOIN " + schema + ".whouse w ON dr.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON dr.supl_code=m7.c_code  LEFT JOIN " + schema + ".m10 m ON dr.pay_code=m.mp_code WHERE dr.inv_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' ORDER BY dr.inv_date");
        }

        public DataTable get_dr_bynumber(String nofrm, String noto)
        {
            return this.QueryBySQLCode("SELECT dr.*, w.*, m7.*, m.* FROM " + schema + ".pinvhd dr LEFT JOIN " + schema + ".whouse w ON dr.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON dr.supl_code=m7.c_code  LEFT JOIN " + schema + ".m10 m ON dr.pay_code=m.mp_code WHERE dr.inv_num BETWEEN '" + nofrm + "' AND '" + noto + "' ORDER BY dr.inv_num");
        }

        public DataTable get_dr_bysupplier(String supfrm, String supto)
        {
            return this.QueryBySQLCode("SELECT dr.*, drl.*, w.*, m7.* FROM " + schema + ".pinvhd dr LEFT JOIN " + schema + ".pinvln drl ON dr.inv_num=drl.inv_num LEFT JOIN " + schema + ".whouse w ON dr.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON dr.supl_code=m7.c_code WHERE dr.supl_code BETWEEN '" + supfrm + "' AND '" + supto + "' ORDER BY dr.supl_code, drl.inv_num");
        }

        public DataTable get_dr_byitems(String itmfrm, String itmto)
        {
            return this.QueryBySQLCode("SELECT dr.*, drl.*, w.*, m7.* FROM " + schema + ".pinvhd dr LEFT JOIN " + schema + ".pinvln drl ON dr.inv_num=drl.inv_num LEFT JOIN " + schema + ".whouse w ON dr.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON dr.supl_code=m7.c_code WHERE drl.item_code BETWEEN '" + itmfrm + "' AND '" + itmfrm + "' ORDER BY drl.item_code, drl.inv_num");
        }

        //stock adjustment 
        public DataTable get_stkadj_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.trnx_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' AND trn_type='A' ORDER BY r.trnx_date");
        }

        public DataTable get_stkadj_bynumber(String nofrm, String noto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.purc_ord BETWEEN '" + nofrm + "' AND '" + noto + "' AND trn_type='A' ORDER BY r.rec_num");
        }

        public DataTable get_stkadj_byitem(String itmfrm, String itmto)
        {
            return this.QueryBySQLCode("SELECT r.*, rl.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE rl.item_code BETWEEN '" + itmfrm + "' AND '" + itmto + "' AND trn_type='A' ORDER BY rl.rec_num, rl.item_code");
        }

        //stock issuance 
        public DataTable get_stkiss_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.trnx_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' AND trn_type='I' ORDER BY r.trnx_date");
        }

        public DataTable get_stkiss_bynumber(String nofrm, String noto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.purc_ord BETWEEN '" + nofrm + "' AND '" + noto + "' AND trn_type='I' ORDER BY r.rec_num");
        }

        public DataTable get_stkiss_byscc(String nofrm, String noto)
        {
            return this.QueryBySQLCode("");
        }

        public DataTable get_stkiss_byitem(String itmfrm, String itmto)
        {
            return this.QueryBySQLCode("SELECT r.*, rl.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE rl.item_code BETWEEN '" + itmfrm + "' AND '" + itmto + "' AND trn_type='I' ORDER BY rl.rec_num, rl.item_code");
        }

        //stock transfer 
        public DataTable get_stktra_bydate(String dtfrm, String dtto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.trnx_date BETWEEN '" + dtfrm + "' AND '" + dtto + "' AND trn_type='T' ORDER BY r.trnx_date");
        }

        public DataTable get_stktra_bynumber(String nofrm, String noto)
        {
            return this.QueryBySQLCode("SELECT r.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE r.purc_ord BETWEEN '" + nofrm + "' AND '" + noto + "' AND trn_type='T' ORDER BY r.rec_num");
        }

        public DataTable get_stktra_byitem(String itmfrm, String itmto)
        {
            return this.QueryBySQLCode("SELECT r.*, rl.*, w.*, m7.* FROM " + schema + ".rechdr r LEFT JOIN " + schema + ".reclne rl ON r.rec_num=rl.rec_num LEFT JOIN " + schema + ".whouse w ON r.whs_code=w.whs_code LEFT JOIN " + schema + ".m07 m7 ON r.supl_code=m7.c_code WHERE rl.item_code BETWEEN '" + itmfrm + "' AND '" + itmto + "' AND trn_type='T' ORDER BY rl.rec_num, rl.item_code");
        }

        public String get_mp_code_of_supplier(String c_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m07", "mp_code", "c_code='" + c_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception) { }

            return val;
        }

        public String get_supplier_name(String c_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m07", "c_name", "c_code='" + c_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception) { }

            return val;
        }

        public String get_supplier_contact_person(String c_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m07", "c_cntc", "c_code='" + c_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception) { }

            return val;
        }

        public String get_supplier_contact_number(String c_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m07", "c_tel", "c_code='" + c_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception) { }

            return val;
        }

        public String get_supplier_fax_number(String c_code)
        {
            DataTable dt = this.QueryOnTableWithParams("m07", "c_fax", "c_code='" + c_code + "'", "");
            String val = "";

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    val = dt.Rows[i][0].ToString();
                }
            }
            catch (Exception) { }

            return val;
        }

        public DataTable get_notification(DateTime dtfrm, DateTime dtto)
        {
            return QueryOnTableWithParams("notification", "*", "t_date BETWEEN '" + dtfrm.ToString("yyyy-MM-dd") + "' AND '" + dtto.ToString("yyyy-MM-dd") + "'", " ORDER BY t_date DESC");
        }
    }
}