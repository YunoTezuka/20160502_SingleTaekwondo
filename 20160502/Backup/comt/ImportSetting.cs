using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//
using System.IO;
using Microsoft.Office.Interop.Excel;

/* 
 * 加入這兩個using
 * 並且加入Microsoft.Office.Interop.Excel和Office這兩個參考
 * 然後要將Microsoft.Office.Interop.Excel參考屬性中的"內嵌Interop型別"改為False
 * */

namespace comt
{
    public partial class ImportSetting : Form
    {
        //string PathFile = Directory.GetCurrentDirectory() + @"\Data.xlsx";
        //變數_選手資料//////////////////////////////////////////////////////
        public static string[,] ImportData;
        public static string Select_Pose;//看選單雙品勢
        public static int Raw_CompetitorNumber;//看有幾個參賽者
        public static int Column_Pose1;//若單品勢共讀出多少行
        public static int Column_Pose2;//若雙品勢共讀出多少行
        public static int Count_Referee;//看有幾個評審
        public static string[,] ImportData_Score;
        //變數_選手資料//////////////////////////////////////////////////////

        //Excel檔案位置//////////////////////////////////////////////////////
        public static string ExcelFileLocation;
        //Excel檔案位置//////////////////////////////////////////////////////
        
        
        public ImportSetting()
        {
            InitializeComponent();
        }

        private void btn_Confirm_Click(object sender, EventArgs e)
        {
            string PathFile = comboBox1.Text;
            //dgv_data.DataSource = GetDataTable(PathFile, "Sheet1");
            _Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook myBook = myExcel.Workbooks.Open(PathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet mySheet = (Worksheet)myBook.Worksheets[1];
            int Width = 0, Height = 0, i = 1, j = 1;
            while (((Range)mySheet.Cells[1, i]).Text.ToString() != "")
            {
                i++;
            }
            Width = i - 1;
            while (((Range)mySheet.Cells[j, 1]).Text.ToString() != "")
            {
                j++;
            }
            Height = j - 1;
            lb_Competitor.Text = (Height - 1).ToString();
            Raw_CompetitorNumber = Height - 1;//選手個數
            if (cb_Pose.Text == "1")
            {
                ImportData = new string[17, (Height)];
                for (i = 1; i <= Height; i++)
                {
                    ImportData[0, i - 1] = ((Range)mySheet.Cells[i, 1]).Text.ToString();//組別
                    ImportData[1, i - 1] = ((Range)mySheet.Cells[i, 2]).Text.ToString();//單位
                    ImportData[2, i - 1] = ((Range)mySheet.Cells[i, 3]).Text.ToString();//姓名
                    ImportData[3, i - 1] = ((Range)mySheet.Cells[i, 4]).Text.ToString();//初賽或複賽
                    ImportData[4, i - 1] = ((Range)mySheet.Cells[i, 5]).Text.ToString();//第一品勢名稱
                    ImportData[5, i - 1] = ((Range)mySheet.Cells[i, 6]).Text.ToString();//第二品勢名稱
                    ImportData[6, i - 1] = ((Range)mySheet.Cells[i, 7]).Text.ToString();//第一品勢正確性分數
                    ImportData[7, i - 1] = ((Range)mySheet.Cells[i, 8]).Text.ToString();//第一品勢表現性分數
                    ImportData[8, i - 1] = ((Range)mySheet.Cells[i, 9]).Text.ToString();//第一品勢平均
                    ImportData[9, i - 1] = ((Range)mySheet.Cells[i, 10]).Text.ToString();//第二品勢正確性分數
                    ImportData[10, i - 1] = ((Range)mySheet.Cells[i, 11]).Text.ToString();//第二品勢表現性分數
                    ImportData[11, i - 1] = ((Range)mySheet.Cells[i, 12]).Text.ToString();//第二品勢平均
                    ImportData[12, i - 1] = ((Range)mySheet.Cells[i, 13]).Text.ToString();//總平均
                    ImportData[13, i - 1] = ((Range)mySheet.Cells[i, 14]).Text.ToString();//排名(數字)
                    ImportData[14, i - 1] = ((Range)mySheet.Cells[i, 15]).Text.ToString();//名次(國字)
                    ImportData[15, i - 1] = ((Range)mySheet.Cells[i, 16]).Text.ToString();//編號
                    ImportData[16, i - 1] = ((Range)mySheet.Cells[i, 17]).Text.ToString();//棄權
                    /*
                    ImportData[0, i - 1] = ((Range)mySheet.Cells[i, 1]).Text.ToString();//組別
                    ImportData[1, i - 1] = ((Range)mySheet.Cells[i, 2]).Text.ToString();//單位
                    ImportData[2, i - 1] = ((Range)mySheet.Cells[i, 3]).Text.ToString();//姓名
                    ImportData[3, i - 1] = ((Range)mySheet.Cells[i, 4]).Text.ToString();//初賽或複賽
                    ImportData[4, i - 1] = ((Range)mySheet.Cells[i, 5]).Text.ToString();//第一品勢名稱
                    ImportData[5, i - 1] = ((Range)mySheet.Cells[i, 7]).Text.ToString();//第一品勢正確性分數
                    ImportData[6, i - 1] = ((Range)mySheet.Cells[i, 8]).Text.ToString();//第一品勢表現性分數
                    ImportData[7, i - 1] = ((Range)mySheet.Cells[i, 9]).Text.ToString();//第一品勢平均
                    ImportData[8, i - 1] = ((Range)mySheet.Cells[i, 13]).Text.ToString();//總平均
                    ImportData[9, i - 1] = ((Range)mySheet.Cells[i, 14]).Text.ToString();//排名(數字)
                    ImportData[10, i - 1] = ((Range)mySheet.Cells[i, 15]).Text.ToString();//名次(國字)
                    ImportData[11, i - 1] = ((Range)mySheet.Cells[i, 16]).Text.ToString();//編號
                    ImportData[12, i - 1] = ((Range)mySheet.Cells[i, 17]).Text.ToString();//棄權
                    */
                }

                gb_Pose2.Visible = true;
                gb_Pose1.Visible = false;

                label11.Visible = true;
                label12.Visible = true;
                label13.Visible = true;
                label14.Visible = true;
                label15.Visible = true;
                label16.Visible = true;
                label17.Visible = true;
                label18.Visible = true;
                label19.Visible = true;
                label20.Visible = true;
                label21.Visible = true;
                label22.Visible = true;
                label23.Visible = true;
                label24.Visible = true;
                label25.Visible = true;
                label28.Visible = true;
                label29.Visible = true;

                label11.Text = ImportData[0, 0];
                label12.Text = ImportData[1, 0];
                label13.Text = ImportData[2, 0];
                label14.Text = ImportData[3, 0];
                label15.Text = ImportData[4, 0];
                label16.Text = ImportData[5, 0];
                label17.Text = ImportData[6, 0];
                label18.Text = ImportData[7, 0];
                label19.Text = ImportData[8, 0];
                label20.Text = ImportData[9, 0];
                label21.Text = ImportData[10, 0];
                label22.Text = ImportData[11, 0];
                label23.Text = ImportData[12, 0];
                label24.Text = ImportData[13, 0];
                label25.Text = ImportData[14, 0];
                label28.Text = ImportData[15, 0];
                label29.Text = ImportData[16, 0];
                /*
                gb_Pose1.Visible = true;
                gb_Pose2.Visible = false;

                label1.Visible = true;
                label2.Visible = true;
                label3.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                label10.Visible = true;
                label26.Visible = true;
                label27.Visible = true;

                label1.Text = ImportData[0, 0];
                label2.Text = ImportData[1, 0];
                label3.Text = ImportData[2, 0];
                label4.Text = ImportData[3, 0];
                label5.Text = ImportData[4, 0];
                label6.Text = ImportData[5, 0];
                label7.Text = ImportData[6, 0];
                label8.Text = ImportData[7, 0];
                label9.Text = ImportData[8, 0];
                label10.Text = ImportData[9, 0];
                label26.Text = ImportData[10, 0];
                label27.Text = ImportData[11, 0];
                */
                mySheet = (Worksheet)myBook.Worksheets[1];

                ImportData_Score = new string[29, (Height)];
                /*
                for (i = 1; i <= Height; i++)
                {
                    ImportData_Score[0, i - 1] = ((Range)mySheet.Cells[i, 1]).Text.ToString();
                    ImportData_Score[1, i - 1] = ((Range)mySheet.Cells[i, 2]).Text.ToString();
                    ImportData_Score[2, i - 1] = ((Range)mySheet.Cells[i, 3]).Text.ToString();
                    ImportData_Score[3, i - 1] = ((Range)mySheet.Cells[i, 4]).Text.ToString();
                    ImportData_Score[4, i - 1] = ((Range)mySheet.Cells[i, 5]).Text.ToString();
                    ImportData_Score[5, i - 1] = ((Range)mySheet.Cells[i, 6]).Text.ToString();
                    ImportData_Score[6, i - 1] = ((Range)mySheet.Cells[i, 7]).Text.ToString();
                    ImportData_Score[7, i - 1] = ((Range)mySheet.Cells[i, 8]).Text.ToString();
                    ImportData_Score[8, i - 1] = ((Range)mySheet.Cells[i, 9]).Text.ToString();
                    ImportData_Score[9, i - 1] = ((Range)mySheet.Cells[i, 10]).Text.ToString();
                    ImportData_Score[10, i - 1] = ((Range)mySheet.Cells[i, 11]).Text.ToString();
                    ImportData_Score[11, i - 1] = ((Range)mySheet.Cells[i, 12]).Text.ToString();
                    ImportData_Score[12, i - 1] = ((Range)mySheet.Cells[i, 13]).Text.ToString();
                    ImportData_Score[13, i - 1] = ((Range)mySheet.Cells[i, 14]).Text.ToString();
                    ImportData_Score[14, i - 1] = ((Range)mySheet.Cells[i, 15]).Text.ToString();
                    ImportData_Score[15, i - 1] = ((Range)mySheet.Cells[i, 16]).Text.ToString();
                    ImportData_Score[16, i - 1] = ((Range)mySheet.Cells[i, 17]).Text.ToString();
                    ImportData_Score[17, i - 1] = ((Range)mySheet.Cells[i, 18]).Text.ToString();
                    ImportData_Score[18, i - 1] = ((Range)mySheet.Cells[i, 19]).Text.ToString();
                    ImportData_Score[19, i - 1] = ((Range)mySheet.Cells[i, 20]).Text.ToString();
                    ImportData_Score[20, i - 1] = ((Range)mySheet.Cells[i, 21]).Text.ToString();
                    ImportData_Score[21, i - 1] = ((Range)mySheet.Cells[i, 22]).Text.ToString();
                    ImportData_Score[22, i - 1] = ((Range)mySheet.Cells[i, 23]).Text.ToString();
                    ImportData_Score[23, i - 1] = ((Range)mySheet.Cells[i, 24]).Text.ToString();
                    ImportData_Score[24, i - 1] = ((Range)mySheet.Cells[i, 25]).Text.ToString();
                    ImportData_Score[25, i - 1] = ((Range)mySheet.Cells[i, 26]).Text.ToString();
                    ImportData_Score[26, i - 1] = ((Range)mySheet.Cells[i, 27]).Text.ToString();
                    ImportData_Score[27, i - 1] = ((Range)mySheet.Cells[i, 28]).Text.ToString();
                    ImportData_Score[28, i - 1] = ((Range)mySheet.Cells[i, 29]).Text.ToString();
                }
                */

                Column_Pose1 = 17;
                Select_Pose = "1";
                //btn_Confirm.Text = ImportData[3, 0];
            }
            else if (cb_Pose.Text == "2")
            {
                ImportData = new string[17, (Height)];
                for (i = 1; i <= Height; i++)
                {
                    ImportData[0, i - 1] = ((Range)mySheet.Cells[i, 1]).Text.ToString();//組別
                    ImportData[1, i - 1] = ((Range)mySheet.Cells[i, 2]).Text.ToString();//單位
                    ImportData[2, i - 1] = ((Range)mySheet.Cells[i, 3]).Text.ToString();//姓名
                    ImportData[3, i - 1] = ((Range)mySheet.Cells[i, 4]).Text.ToString();//初賽或複賽
                    ImportData[4, i - 1] = ((Range)mySheet.Cells[i, 5]).Text.ToString();//第一品勢名稱
                    ImportData[5, i - 1] = ((Range)mySheet.Cells[i, 6]).Text.ToString();//第二品勢名稱
                    ImportData[6, i - 1] = ((Range)mySheet.Cells[i, 7]).Text.ToString();//第一品勢正確性分數
                    ImportData[7, i - 1] = ((Range)mySheet.Cells[i, 8]).Text.ToString();//第一品勢表現性分數
                    ImportData[8, i - 1] = ((Range)mySheet.Cells[i, 9]).Text.ToString();//第一品勢平均
                    ImportData[9, i - 1] = ((Range)mySheet.Cells[i, 10]).Text.ToString();//第二品勢正確性分數
                    ImportData[10, i - 1] = ((Range)mySheet.Cells[i, 11]).Text.ToString();//第二品勢表現性分數
                    ImportData[11, i - 1] = ((Range)mySheet.Cells[i, 12]).Text.ToString();//第二品勢平均
                    ImportData[12, i - 1] = ((Range)mySheet.Cells[i, 13]).Text.ToString();//總平均
                    ImportData[13, i - 1] = ((Range)mySheet.Cells[i, 14]).Text.ToString();//排名(數字)
                    ImportData[14, i - 1] = ((Range)mySheet.Cells[i, 15]).Text.ToString();//名次(國字)
                    ImportData[15, i - 1] = ((Range)mySheet.Cells[i, 16]).Text.ToString();//編號
                    ImportData[16, i - 1] = ((Range)mySheet.Cells[i, 17]).Text.ToString();//棄權
                }

                gb_Pose2.Visible = true;
                gb_Pose1.Visible = false;

                label11.Visible = true;
                label12.Visible = true;
                label13.Visible = true;
                label14.Visible = true;
                label15.Visible = true;
                label16.Visible = true;
                label17.Visible = true;
                label18.Visible = true;
                label19.Visible = true;
                label20.Visible = true;
                label21.Visible = true;
                label22.Visible = true;
                label23.Visible = true;
                label24.Visible = true;
                label25.Visible = true;
                label28.Visible = true;
                label29.Visible = true;

                label11.Text = ImportData[0, 0];
                label12.Text = ImportData[1, 0];
                label13.Text = ImportData[2, 0];
                label14.Text = ImportData[3, 0];
                label15.Text = ImportData[4, 0];
                label16.Text = ImportData[5, 0];
                label17.Text = ImportData[6, 0];
                label18.Text = ImportData[7, 0];
                label19.Text = ImportData[8, 0];
                label20.Text = ImportData[9, 0];
                label21.Text = ImportData[10, 0];
                label22.Text = ImportData[11, 0];
                label23.Text = ImportData[12, 0];
                label24.Text = ImportData[13, 0];
                label25.Text = ImportData[14, 0];
                label28.Text = ImportData[15, 0];
                label29.Text = ImportData[16, 0];


                mySheet = (Worksheet)myBook.Worksheets[1];

                ImportData_Score = new string[29, (Height)];
                /*
                for (i = 1; i <= Height; i++)
                {
                    ImportData_Score[0, i - 1] = ((Range)mySheet.Cells[i, 1]).Text.ToString();
                    ImportData_Score[1, i - 1] = ((Range)mySheet.Cells[i, 2]).Text.ToString();
                    ImportData_Score[2, i - 1] = ((Range)mySheet.Cells[i, 3]).Text.ToString();
                    ImportData_Score[3, i - 1] = ((Range)mySheet.Cells[i, 4]).Text.ToString();
                    ImportData_Score[4, i - 1] = ((Range)mySheet.Cells[i, 5]).Text.ToString();
                    ImportData_Score[5, i - 1] = ((Range)mySheet.Cells[i, 6]).Text.ToString();
                    ImportData_Score[6, i - 1] = ((Range)mySheet.Cells[i, 7]).Text.ToString();
                    ImportData_Score[7, i - 1] = ((Range)mySheet.Cells[i, 8]).Text.ToString();
                    ImportData_Score[8, i - 1] = ((Range)mySheet.Cells[i, 9]).Text.ToString();
                    ImportData_Score[9, i - 1] = ((Range)mySheet.Cells[i, 10]).Text.ToString();
                    ImportData_Score[10, i - 1] = ((Range)mySheet.Cells[i, 11]).Text.ToString();
                    ImportData_Score[11, i - 1] = ((Range)mySheet.Cells[i, 12]).Text.ToString();
                    ImportData_Score[12, i - 1] = ((Range)mySheet.Cells[i, 13]).Text.ToString();
                    ImportData_Score[13, i - 1] = ((Range)mySheet.Cells[i, 14]).Text.ToString();
                    ImportData_Score[14, i - 1] = ((Range)mySheet.Cells[i, 15]).Text.ToString();
                    ImportData_Score[15, i - 1] = ((Range)mySheet.Cells[i, 16]).Text.ToString();
                    ImportData_Score[16, i - 1] = ((Range)mySheet.Cells[i, 17]).Text.ToString();
                    ImportData_Score[17, i - 1] = ((Range)mySheet.Cells[i, 18]).Text.ToString();
                    ImportData_Score[18, i - 1] = ((Range)mySheet.Cells[i, 19]).Text.ToString();
                    ImportData_Score[19, i - 1] = ((Range)mySheet.Cells[i, 20]).Text.ToString();
                    ImportData_Score[20, i - 1] = ((Range)mySheet.Cells[i, 21]).Text.ToString();
                    ImportData_Score[21, i - 1] = ((Range)mySheet.Cells[i, 22]).Text.ToString();
                    ImportData_Score[22, i - 1] = ((Range)mySheet.Cells[i, 23]).Text.ToString();
                    ImportData_Score[23, i - 1] = ((Range)mySheet.Cells[i, 24]).Text.ToString();
                    ImportData_Score[24, i - 1] = ((Range)mySheet.Cells[i, 25]).Text.ToString();
                    ImportData_Score[25, i - 1] = ((Range)mySheet.Cells[i, 26]).Text.ToString();
                    ImportData_Score[26, i - 1] = ((Range)mySheet.Cells[i, 27]).Text.ToString();
                    ImportData_Score[27, i - 1] = ((Range)mySheet.Cells[i, 28]).Text.ToString();
                    ImportData_Score[28, i - 1] = ((Range)mySheet.Cells[i, 29]).Text.ToString();
                }
                */
                
                Column_Pose1 = 17;
                Select_Pose = "2";
            }

            Count_Referee = int.Parse(cb_RefereeNumber.Text);//有多少評審

            myBook.Save();
            myBook.Close(false, Type.Missing, Type.Missing);//關閉Excel
            myExcel.Quit();//釋放Excel資源            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            myExcel = null;
            mySheet = null;
            //Range = null;
            myBook = null;
            GC.Collect();

            this.Close();
        }

        private void btn_Browse_Click(object sender, EventArgs e)
        {
            //Directory.GetCurrentDirectory() + @"\Data.xlsx";
            openFileDialog1.ShowDialog();
            comboBox1.Text = openFileDialog1.FileName;
            ExcelFileLocation = comboBox1.Text;
        }
        /*
        public static System.Data.DataTable GetDataTable(string FileName, string SheetName)
        {
            System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection
                ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties= Excel 8.0;");

            connection.Open();

            string query = "select * from [" + SheetName + "$]";

            System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(query,  connection);
            System.Data.DataSet ds = new System.Data.DataSet();

            adapter.Fill(ds);

            return ds.Tables[0];
        }
        */

        
    }
}
