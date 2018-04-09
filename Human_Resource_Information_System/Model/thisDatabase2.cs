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
    public class thisDatabase2
    {
        public static String server = "localhost";// System.IO.File.ReadAllText(thisDatabase.driveloc + thisDatabase.comp_folder + "\\HRPublish\\serverDatabase.txt"); //
        public static String lcl_db = thisDatabase.lcl_db;
        public static String svr_pass = thisDatabase.svr_pass;
        String schema = thisDatabase.schema_static;//*/

        NpgsqlConnection conn = new NpgsqlConnection("Server=" + server + ";Port=5432;User Id=postgres;Password=" + svr_pass + ";Database=" + lcl_db + ";");

        GlobalMethod gm = new GlobalMethod();

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

        public Boolean InsertOnTable(String table, String column, String value)
        {
            Boolean flag = false;

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

            try
            {
                this.OpenConn();

                if (cond != "")
                {
                    cond = " WHERE " + cond + "";
                }

                string SQL = "UPDATE " + this.schema + "." + table + " SET " + col_upd + "" + cond + ";";
                //MessageBox.Show(SQL);
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

                if (cond == null || cond == "")
                {
                    cond = "";
                }
                else
                {
                    cond = " WHERE " + cond;
                }

                string SQL = "SELECT " + param + " FROM " + this.schema + "." + table + " " + cond + " " + addcode;
                //MessageBox.Show(SQL);
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
    }
}