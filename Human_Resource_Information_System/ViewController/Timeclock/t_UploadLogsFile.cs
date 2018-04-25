using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace Human_Resource_Information_System
{
    public partial class t_UploadLogsFile : Form
    {

        private GlobalClass gc;
        private GlobalMethod gm;
        thisDatabase db = new thisDatabase();

        public t_UploadLogsFile()
        {
            InitializeComponent();
        }

        private void t_UploadLogsFile_Load(object sender, EventArgs e)
        {
            dips_list();
            pbar.Hide();
        }

        private void dgv_list_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        void dips_list() {

            DataTable dt;
            dt = db.QueryBySQLCode("SELECT * FROM rssys.hr_tito2 ORDER BY work_date DESC");
            dgv_list.DataSource = dt;
        }
        private void btn_browse_Click_1(object sender, EventArgs e)
        {
            
            OpenFileDialog fDialog = new OpenFileDialog();
            fDialog.Title = "Select file to be upload";
            fDialog.Filter = "(*.txt)|*.txt";
            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fDialog.FileName.ToString();
                   
            }

        }

        private void btn_upload_Click(object sender, EventArgs e)
        {


            int rCnt = 0;
            int rw = 0;
            int cl = 0;
            String filename = "";
            String col = "", val = "";
            DataTable data;
            String line = "";
            String table = "hr_tito2";
            String empid = "", logs_id="",time_log = "", status = "", source = "",dt = "",c_tlog="",in_out="",staticval="";
            DateTime work_date;
            //DateTime excel_time;

            
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Please select a file to be uploaded.");
            }
            else
            {
                      
                String pattern = "\\s+";
                String replacement = " ";
                Regex rgx = new Regex(pattern);
                String input = textBox1.Text;
                StreamReader sr = new StreamReader(textBox1.Text);
                DataTable bio_empid = null;
                String bio_id = "";
                string temp = "";
                int row_line = 0;
                while(line !=null)
                {
                    data = null;
                    line= sr.ReadLine();
                    if(line!=null)
                    {
                       try
                       {
                            String yow = line;
                            string result = rgx.Replace(yow, replacement);
                            string[] split = result.Split(' ');

                            bio_id = split[0]; //get employee id WHERE biometric='split[0]'

                            bio_empid = db.QueryBySQLCode("SELECT empid FROM rssys.hr_employee WHERE biometric = '" + bio_id + "' LIMIT 1");
                            if (bio_empid != null && bio_empid.Rows.Count > 0)
                            {
                                empid = bio_empid.Rows[0]["empid"].ToString();
                                work_date = Convert.ToDateTime(split[1]);

                                time_log = temp = split[2];
                                in_out = split[3];
                                staticval = split[4];

                               // MessageBox.Show("Empid : " + empid + " Workdate : " + work_date + " Line : " + row_line + "Time Log : " + time_log);
                                pbar.Show();
                                if (in_out == staticval)
                                {
                                    status = "O";
                                }
                                else
                                {
                                    status = "I";
                                }
                                source = "M";

                                
                                data = db.QueryBySQLCode("SELECT * FROM rssys.hr_tito2 WHERE empid = '" + empid + "' AND work_date='" + work_date.ToString("yyyy-MM-dd") + "' AND time_log='" + time_log + "' AND status = '" + status +"'");
                                if (data != null && data.Rows.Count <= 0)
                                {
                                    logs_id = db.get_pk("logs_id");
                                    col = "logs_id,work_date,time_log,empid,status,source";
                                    val = "'" + logs_id + "','" + work_date.ToString("yyyy-MM-dd") + "','" + time_log + "','" + empid + "','" + status + "','" + source + "'";

                                    db.InsertOnTable(table, col, val);
                                    db.set_pkm99("logs_id", db.get_nextincrementlimitchar(logs_id, 8));
                                    data = null;
                                }
                                
                                if (rCnt != 100 || rCnt < 100)
                                {
                                    pbar.Value = rCnt++;
                                }
                            }
                            else { continue; }
                            
                        }
                        catch (Exception er)
                        {
                            MessageBox.Show(er.StackTrace + "\n : Bio ID :" + bio_id + " Temp :" + temp + "Empid : " +empid  + "Row : " + row_line);
                        }
                    }
                    row_line++;
                }
                sr.Close();
                DialogResult results = MessageBox.Show("File Uploaded", "Confirmation", MessageBoxButtons.OK);
                if (results == DialogResult.OK)
                {
                    pbar.Value = 0;
                    pbar.Hide();
                    textBox1.Text = "";
                }
            }

            dips_list();
                
        }


        public String getTimeString(Excel.Range range, int row, int col)
        {
            DateTime dt = DateTime.Now;
            String dtstr = "";
            if (range != null)
            {
                try
                {
                    dtstr = getString(range, row, col);
                    try { dt = DateTime.Parse(dt.ToString("yyyy-MM-dd ") + dtstr); }
                    catch { dt = DateTime.FromOADate(Double.Parse(dtstr)); }
                }
                catch { }
            }
            return dt.ToString("HH:mm");
        }
        public String getString(Excel.Range range, int row, int col)
        {
            String str = "";
            if (range != null)
            {
                try
                {
                    str = Convert.ToString((range.Cells[row, col] as Excel.Range).Value2 ?? "");
                }
                catch { }
            }
            return str;
        }   
        public String getDateString(Excel.Range range, int row, int col)
        {
            DateTime dt = DateTime.Now;
            String dtstr = "";
            if (range != null)
            {
                try
                {
                    dtstr = getString(range, row, col);
                    try { dt = DateTime.Parse(dtstr); }
                    catch { dt = DateTime.FromOADate(Double.Parse(dtstr)); }
                }
                catch { }
            }
            return dt.ToString("yyyy-MM-dd");
        }
        private void dgv_list_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
