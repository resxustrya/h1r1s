using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Human_Resource_Information_System
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Login());
            try
            {
                Application.SetCompatibleTextRenderingDefault(false);
                Application.EnableVisualStyles();
                DialogResult result;


                using (var loginForm = new Login())

                    result = loginForm.ShowDialog();

                if (result == DialogResult.OK)
                {
                    // login was successful
                    Application.Run(new Main());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Unexpected exception");
            }
           
        }


    }
}
