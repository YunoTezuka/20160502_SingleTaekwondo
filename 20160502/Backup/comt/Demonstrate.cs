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

        public Demonstrate()
        {
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
                lb_Time.Text = "";
                timer_Go.Enabled = true;
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
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

                lb_Pose1.Text = comt.Form1.ImportData[4, comt.Form1.Now_Player];//第一品勢名稱
                lb_Pose2.Text = comt.Form1.ImportData[5, comt.Form1.Now_Player];//第二品勢名稱
                lb_Time.Text = "";
                timer_Go.Enabled = true;
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
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
                 this.lb_Time.Invoke(
                 new MethodInvoker(
                 delegate
                 {
                     Timer_Count++;
                     String minute = (Timer_Count / 60).ToString().PadLeft(2, '0');
                     String second = (Timer_Count % 60).ToString().PadLeft(2, '0');
                     this.lb_Time.Text = minute + ":" + second;
                 }
                 )
                );
             }
                
         }
         
         private void Demonstrate_FormClosing(object sender, FormClosingEventArgs e)
         {
             timer.Enabled = false;             
         }

    }
}
