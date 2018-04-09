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
        void dips_list()
        {

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

            //Excel.Application xlApp;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //Excel.Range range;


            int rCnt = 0;
            int rw = 0;
            int cl = 0;
            String filename = "";
            String col = "", val = "";
            DataTable data;
            String line = "";
            String table = "hr_tito2";
            String empid = "", work_date = "", time_log = "", status = "", source = "", dt = "", c_tlog = "", in_out = "", staticval = "";
            String bio = "";
            DataTable dtcheck;
            //DateTime excel_time;


            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Please select an excel file to import.");
            }
            else
            {
                try
                {
                    //filename = textBox1.Text;
                    //xlApp = new Excel.Application();
                    //xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    //range = xlWorkSheet.UsedRange;
                    //rw = range.Rows.Count;
                    //cl = range.Columns.Count;
                    //pbar.Maximum = rw;
                    String pattern = "\\s+";
                    String replacement = " ";
                    Regex rgx = new Regex(pattern);
                    String input = textBox1.Text;
                    StreamReader sr = new StreamReader(textBox1.Text);
                    while (line != null)
                    {

                        line = sr.ReadLine();
                        if (line != null)
                        {
                            String yow = line;
                            string result = rgx.Replace(yow, replacement);

                            string[] split = result.Split(' ');
                            bio = split[0];
                            work_date = DateTime.Parse(split[1]).ToString("yyyy-MM-dd");
                            time_log = split[2];
                            in_out = split[3];
                            staticval = split[4];

                            pbar.Show();
                            dtcheck = db.QueryBySQLCode("SELECT empid from rssys.hr_employee WHERE biometric='" + bio + "'");

                            if (dtcheck.Rows.Count > 0)
                            {
                                empid = dtcheck.Rows[0]["empid"].ToString();

                                //for (rCnt = 2; rCnt <= rw; rCnt++)
                                //{


                                //    if (rCnt != 1)
                                //    {
                                //empid = getString(range, rCnt, 1);
                                //work_date = getDateString(range, rCnt, 2);
                                //time_log = getTimeString(range, rCnt, 3);
                                //in_out = getString(range, rCnt, 4);
                                //staticval = getString(range, rCnt, 5);

                                //empid = "";
                                //work_date = "";
                                //time_log = "";
                                //in_out = "";
                                //staticval = "";

                                if (in_out == staticval)
                                {
                                    status = "O";
                                }
                                else
                                {
                                    status = "I";
                                }
                                source = "M";


                                col = "work_date,time_log,empid,status,source";
                                val = "'" + work_date + "','" + time_log + "','" + empid + "','" + status + "','" + source + "'";


                                data = db.QueryBySQLCode("SELECT * FROM rssys.hr_tito2 WHERE empid = '" + empid + "' AND work_date='" + work_date + "' AND time_log='" + time_log + "'");
                                if (data.Rows.Count < 1)
                                {
                                    data = null;
                                    data = db.QueryBySQLCode("SELECT * FROM rssys.hr_tito2 WHERE empid = '" + empid + "' AND work_date='" + work_date + "' AND status='" + status + "'");

                                    if (data.Rows.Count > 0)
                                    {



                                    }
                                    else
                                    {

                                        db.InsertOnTable(table, col, val);
                                        data = null;

                                    }

                                }






                                //}


                                if (rCnt != 100 || rCnt < 100)
                                {
                                    pbar.Value = rCnt++;
                                }

                                //}

                            }



                        }
                        else
                        {
                            //MessageBox.Show("File is Empty");

                        }


                    }



                    sr.Close();
                    DialogResult results = MessageBox.Show("File Uploaded", "Confirmation", MessageBoxButtons.OK);
                    if (results == DialogResult.OK)
                    {
                        pbar.Value = 0;
                        pbar.Hide();
                        textBox1.Text = "";
                    }

                    if (rw > 0)
                    {


                    }


                }
                catch (Exception ex)
                {
                    DialogResult result = MessageBox.Show(ex.Message, "Confirmation", MessageBoxButtons.OK);
                    if (result == DialogResult.OK)
                    {
                        pbar.Value = 0;
                        pbar.Hide();
                        textBox1.Text = "";
                    }
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
