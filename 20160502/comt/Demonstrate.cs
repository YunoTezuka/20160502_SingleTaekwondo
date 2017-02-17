using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//----------------------------
using System.Timers;
using System.IO;


namespace comt
{
    public partial class Demonstrate : Form
    {

        //變數/////////////////////////////////////////////////
        string[,] ImportData;
        string Select_Pose;
        int Raw_CompetitorNumber;
        int Column_Pose1;
        int Column_Pose2;
        //變數/////////////////////////////////////////////////

        //時間/////////////////////////////////////////////////
        int Timer_Count = -1;
        public static System.Timers.Timer timer;//正數計時用
        //時間/////////////////////////////////////////////////

        //現行狀態State/////////////////////////////////////////////
        int Now_Player;
        int Now_Pose;
        //現行狀態State/////////////////////////////////////////////

        //RemoteControl_Timer(開啟Demonstrate的Timer)////////////////////////////////////
        bool RemoteControl_TimerStart = false;
        //RemoteControl_Timer(開啟Demonstrate的Timer)////////////////////////////////////
        string Country;
        public Demonstrate()
        {

            //移動視窗到延伸桌面
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            InitializeComponent();

            timer = new System.Timers.Timer();
            timer.Elapsed += new ElapsedEventHandler(DisplayTimeEvent);
            timer.Interval = 1000;
            timer.Start();
            timer.Enabled = true;

            Now_Pose = comt.Form1.Now_Pose;

            if (comt.Form1.Now_Pose.ToString() == "1")
            {
                
                pnl_Pose01.BackColor = Color.Yellow;
                pnl_Pose02.BackColor = Color.Gray;
                lb_Pose1.ForeColor = Color.Red;
                lb_Pose2.ForeColor = Color.DarkGray;   
                
                //------------------
                lb_Pose1.Visible = true;
                lb_Pose2.Visible = true;
                lb_Pose_Title.Text = "第一品勢名稱:";
                lb_ID.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名

                lb_Pose1.Text = comt.Form1.ImportData[4, comt.Form1.Now_Player];//第一品勢名稱
                lb_Pose2.Text = comt.Form1.ImportData[5, comt.Form1.Now_Player];//第二品勢名稱
                /*
                lb_ID.Text = comt.Form1.ImportData[11, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名
                lb_Pose1.Text = comt.Form1.ImportData[4, comt.Form1.Now_Player];//第一品勢名稱
                lb_Pose2.Text = comt.Form1.ImportData[5, comt.Form1.Now_Player];//第二品勢名稱
                */
                //lb_Time.Text = "";
                timer_Go.Enabled = true;
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
                lb_GroupNumber.Text = comt.Form1.ImportData[17, comt.Form1.Now_Player];//組別號碼
                //===================================
                //lb_Pose2.Visible = false;
                //===================================  
                
            }
            if (comt.Form1.Now_Pose.ToString() == "2")
            {
               
                pnl_Pose01.BackColor = Color.Gray;
                pnl_Pose02.BackColor = Color.Yellow;
                lb_Pose1.ForeColor = Color.DarkGray;
                lb_Pose2.ForeColor = Color.Red;
                
                //------------------
                lb_Pose2.Visible = true;
                lb_Pose1.Visible = true;
                lb_Pose_Title.Text = "第二品勢名稱:";
                lb_ID.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名
                string temp_jihl = comt.Form1.ImportData[5, comt.Form1.Now_Player]; ;
                lb_Pose1.Text = comt.Form1.ImportData[4, comt.Form1.Now_Player];//第一品勢名稱
                lb_Pose2.Text = comt.Form1.ImportData[5, comt.Form1.Now_Player];//第二品勢名稱
                //lb_Time.Text = "";
                timer_Go.Enabled = true;
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
                lb_GroupNumber.Text = comt.Form1.ImportData[17, comt.Form1.Now_Player];//組別號碼
                //===================================
                //lb_Pose1.Visible = false;
                //===================================
                Country = comt.Form1.ImportData[18, comt.Form1.Now_Player];//國旗
                if (System.IO.File.Exists(comt.Form1.ParentPath + @"\Country\" + Country + ".jpg"))//判斷該檔案是否存在//(Application.StartupPath + @"\Country\" + Country + ".jpg")
                {
                    this.pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
                    Bitmap buffer = new Bitmap(comt.Form1.ParentPath + @"\Country\" + Country + ".jpg");
                    //pb_Country.BackgroundImage = Image.FromFile(Application.StartupPath + @"\Country\" + Country + ".jpg");
                    this.pb_Country.Image = buffer;
                    this.pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
                }
                else
                {
                    pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
                    pb_Country.BackgroundImage = Image.FromFile(comt.Form1.ParentPath + @"\Country\Blank.jpg");
                    
                }
            }
            
            //setFont////////////////////////////////////////////////////////////
            lb_Unit.Font = new Font(lb_Unit.Font.FontFamily, 100.0f, lb_Unit.Font.Style);
            while (lb_Unit.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Unit.Text,
            new Font(lb_Unit.Font.FontFamily, lb_Unit.Font.Size, lb_Unit.Font.Style)).Width)
            {
                lb_Unit.Font = new Font(lb_Unit.Font.FontFamily, lb_Unit.Font.Size - 0.5f, lb_Unit.Font.Style);
            }

            lb_Name.Font = new Font(lb_Name.Font.FontFamily, 100.0f, lb_Name.Font.Style);
            while (lb_Name.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name.Text,
            new Font(lb_Name.Font.FontFamily, lb_Name.Font.Size, lb_Name.Font.Style)).Width)
            {
                lb_Name.Font = new Font(lb_Name.Font.FontFamily, lb_Name.Font.Size - 0.5f, lb_Name.Font.Style);
            }

            lb_Group.Font = new Font(lb_Group.Font.FontFamily, 100.0f, lb_Group.Font.Style);
            while (lb_Group.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Group.Text,
            new Font(lb_Group.Font.FontFamily, lb_Group.Font.Size, lb_Group.Font.Style)).Width)
            {
                lb_Group.Font = new Font(lb_Group.Font.FontFamily, lb_Group.Font.Size - 0.5f, lb_Group.Font.Style);
            }

            
            lb_ID.Font = new Font(lb_ID.Font.FontFamily, 90.0f, lb_ID.Font.Style);
            while (lb_ID.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID.Text,
            new Font(lb_ID.Font.FontFamily, lb_ID.Font.Size, lb_ID.Font.Style)).Width)
            {
                lb_ID.Font = new Font(lb_ID.Font.FontFamily, lb_ID.Font.Size - 0.5f, lb_ID.Font.Style);
            }

            lb_Pose1.Font = new Font(lb_Pose1.Font.FontFamily, 86.0f, lb_Pose1.Font.Style);
            while (lb_Pose1.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Pose1.Text,
            new Font(lb_Pose1.Font.FontFamily, lb_Pose1.Font.Size, lb_Pose1.Font.Style)).Width)
            {
                lb_Pose1.Font = new Font(lb_Pose1.Font.FontFamily, lb_Pose1.Font.Size - 0.5f, lb_Pose1.Font.Style);
            }

            lb_Pose2.Font = new Font(lb_Pose2.Font.FontFamily, 86.0f, lb_Pose2.Font.Style);
            while (lb_Pose2.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Pose2.Text,
            new Font(lb_Pose2.Font.FontFamily, lb_Pose2.Font.Size, lb_Pose2.Font.Style)).Width)
            {
                lb_Pose2.Font = new Font(lb_Pose2.Font.FontFamily, lb_Pose2.Font.Size - 0.5f, lb_Pose2.Font.Style);
            }
            //lb_GroupNumber.Text = "1";
            lb_GroupNumber.Font = new Font(lb_GroupNumber.Font.FontFamily, 100.0f, lb_GroupNumber.Font.Style);
            while (lb_GroupNumber.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_GroupNumber.Text,
            new Font(lb_GroupNumber.Font.FontFamily, lb_GroupNumber.Font.Size, lb_GroupNumber.Font.Style)).Width)
            {
                lb_GroupNumber.Font = new Font(lb_GroupNumber.Font.FontFamily, lb_GroupNumber.Font.Size - 0.5f, lb_GroupNumber.Font.Style);
            }
            

            //setFont////////////////////////////////////////////////////////////
            Country = comt.Form1.ImportData[18, comt.Form1.Now_Player];//國旗
            if (System.IO.File.Exists(comt.Form1.ParentPath + @"\Country\" + Country + ".jpg"))//判斷該檔案是否存在
            {
                //pb_Country.BackgroundImage = Image.FromFile(Application.StartupPath + @"\Country\" + Country + ".jpg");
                pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
                Bitmap buffer = new Bitmap(comt.Form1.ParentPath + @"\Country\" + Country + ".jpg");
                //放置您所指定的圖片
                //並指定圖片要放置的位置，(X,Y) = (0,0)
                //pictureBox2.Width = 500;
                //pictureBox2.Height = 125;
                pb_Country.Image = buffer;
                pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            else
            {
                pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
                Bitmap buffer = new Bitmap(comt.Form1.ParentPath + @"\Country\Blank.jpg");
                pb_Country.Image = buffer;
                //pb_Country.BackgroundImage = Image.FromFile(Application.StartupPath + @"\Country\Blank.jpg");
                pb_Country.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void timer_Go_Tick(object sender, EventArgs e)
        {

        }
         public void DisplayTimeEvent( object source, ElapsedEventArgs e )
         {
             RemoteControl_TimerStart = comt.Form1.RemoteControl_TimerStart;//這個Timer的開啟由Form1 的 RemoteControl_TimerStart這個變數控制

             if (RemoteControl_TimerStart == true) 
             {
                 if (comt.Form1.IsOpen_Demonstrate == true)//用這道判斷式的用意是，為了判斷若invoke出去做事但視窗已經被關閉了那就無法使用label就會當掉
                 {
                     try 
                     {
                         lb_Time.ForeColor = Color.Yellow;//開始才讓字變黃色
                         this.lb_Time.Invoke(
                         new MethodInvoker(
                         delegate
                         {
                             Timer_Count++;
                             String minute = (Timer_Count / 60).ToString();
                             //String minute = (Timer_Count / 60).ToString().PadLeft(2, '0');
                             String second = (Timer_Count % 60).ToString().PadLeft(2, '0');
                             this.lb_Time.Text = minute + ":" + second;
                         }
                         )
                        );
                     }
                     catch(Exception ex)
                     {
                     
                     }                    
                 }
             }
                
         }
         
         private void Demonstrate_FormClosing(object sender, FormClosingEventArgs e)
         {
             timer.Enabled = false;             
         }

    }
}
