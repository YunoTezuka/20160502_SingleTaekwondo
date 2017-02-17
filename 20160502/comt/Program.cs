using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace comt
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try 
            {
                Application.Run(new Form1());
            }
            catch(Exception ex)
            {
                MessageBox.Show("==============================\n原因:" + ex.Message + "\n==============================\n" + ex.Source + "\n==============================\n" + ex.StackTrace);
                if (ex.Message == "通訊埠已經關閉。")
                {
                    Environment.Exit(Environment.ExitCode);
                }
                else 
                { 
                }
            }
            
        }
    }
}
