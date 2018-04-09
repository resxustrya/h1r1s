using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace Human_Resource_Information_System
{
    class GlobalMethod
    {
        public String get_amount_negbracket(Double amnt)
        {
            String new_amnt = "0.00";

            if (amnt < 0)
            {
                new_amnt = "(" + Math.Abs(amnt).ToString("0.00") + ")";
            }
            else
            {
                new_amnt = amnt.ToString("0.00");
            }

            return new_amnt;
        }

        public Boolean is_Doublevalid(String amnt)
        {
            try
            {
                Convert.ToDouble(amnt);

                return true;
            }
            catch (Exception) { }

            return false;
        }

        public String toAccountingFormat(Double amt)
        {
            try
            {
                return amt.ToString("#,##0.00;(#,##0.00);0.00");
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return "";
        }

        public String toAccountingFormat(String amt)
        {
            try
            {
                Double dbl = toNormalDoubleFormat(amt);

                return dbl.ToString("#,##0.00;(#,##0.00);0.00");
            }
            catch (Exception er) { MessageBox.Show(er.Message); }

            return "";
        }

        public Double toNormalDoubleFormat(String acct_amt)
        {
            try
            {
                if (acct_amt.Contains("(") && acct_amt.Contains(")"))
                {
                    return Double.Parse(acct_amt, NumberStyles.AllowParentheses |
                                      NumberStyles.AllowThousands |
                                      NumberStyles.AllowDecimalPoint);
                }
                else
                    return Convert.ToDouble(acct_amt);
            }
            catch (Exception) { }

            return 0.00;
        }

        public String toDoubleStr(String acct_amt)
        {
            try
            {
                if (acct_amt.Contains("(") && acct_amt.Contains(")"))
                {
                    return Double.Parse(acct_amt, NumberStyles.AllowParentheses |
                                      NumberStyles.AllowThousands |
                                      NumberStyles.AllowDecimalPoint).ToString("0.00");
                }
                else
                    return Convert.ToDouble(acct_amt).ToString("0.00");
            }
            catch (Exception) { }

            return "0.00";
        }

        public Int32 toInt(String val)
        {
            int covertval = 0;

            try
            {
                covertval = Convert.ToInt32(val);
            }
            catch (Exception) { }

            return covertval;
        }

        public String toDateString(String date_value, String format)
        {
            String formatteddate = "";

            if (String.IsNullOrEmpty(format))
            {
                format = "yyyy-MM-dd";
            }

            try
            {
                formatteddate = Convert.ToDateTime(date_value).ToString(format);
            }
            catch (Exception) {  }

            return formatteddate;
        }

        public DateTime toDateValue(String date_value)
        {
            DateTime formatteddate = new DateTime();

            try
            {
                formatteddate = Convert.ToDateTime(date_value);
            }
            catch (Exception) { }

            return formatteddate;
        }

    }
}
