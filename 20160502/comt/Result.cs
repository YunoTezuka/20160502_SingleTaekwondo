using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace comt
{
    public partial class Result : Form
    {
        //選手資料///////////////////////////////////////////////////////
        string[,] ImportData;
        string Select_Pose;//共有多少品勢
        int Raw_CompetitorNumber;//有多少位參賽者
        int Column_Pose1;//單品勢有多少欄位
        int Column_Pose2;//雙品勢有多少欄位
        string[,] ImportData_Score;
        //--------------------------------        
        int Count_Referee;//看有幾個評審
        //選手資料///////////////////////////////////////////////////////

        //現行狀態State//////////////////////////////////////////////////////////
        int Now_Player;
        int Now_Pose;
        //現行狀態State//////////////////////////////////////////////////////////


        public Result()
        {
            //SetStyle(ControlStyles.UserPaint, true);
            //SetStyle(ControlStyles.AllPaintingInWmPaint, true); // 禁止擦除背景.
            //SetStyle(ControlStyles.OptimizedDoubleBuffer, true); // 双缓冲
            //移動視窗到延伸桌面
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            InitializeComponent();
            comt.Form1.IsWritng_Rank = true;//正在寫入
            comt.Form1.WriteExcelAll();

            //ini///////////////////////////////////////////////////////////////////
            ImportData = comt.Form1.ImportData;//把exl資料表格載入
            Select_Pose = comt.Form1.Select_Pose;//在ImportSetting時選的要比幾個品勢
            Raw_CompetitorNumber = comt.Form1.Raw_CompetitorNumber;//共有多少位參賽者
            Column_Pose1 = comt.Form1.Column_Pose1;//若選單品勢會有幾行
            Column_Pose2 = comt.Form1.Column_Pose2;//若選雙品勢會有幾行
            Count_Referee = comt.Form1.Count_Referee;//有幾位評審
            ImportData_Score = comt.Form1.ImportData_Score;//把exl分數表格載入    
            //ini///////////////////////////////////////////////////////////////////
            Now_Pose = comt.Form1.Now_Pose;
            Now_Player = comt.Form1.Now_Player;

            if (comt.Form1.Now_Pose.ToString() == "1")
            {
                lb_ID.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名  
                lb_Number.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                //lb_Pose.Text = "第一品勢";
                //-----------------------------------------------------------------------------------
                //Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
            }
            if (comt.Form1.Now_Pose.ToString() == "2")
            {
                lb_ID.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名
                lb_Number.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                //lb_Pose.Text = "第二品勢";
                //-----------------------------------------------------------------------------------
                //Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
            }

            if (Select_Pose == "1" && Now_Pose == 1)
            {   
                //6:1st正當性平均分數 
                //7:1st表現性平均分數 
                //8:第一品勢總分 
                //9:2nd正當性平均分數
                //10:2nd表現性平均分數
                //11:第二品勢總分 
                lb_Pose1_Totle.Text = comt.Form1.ImportData[8, comt.Form1.Now_Player];//第一品勢總分 = 1st正當性+1st表現性總分
                lb_Correctness_Average.Text = comt.Form1.ImportData[6, comt.Form1.Now_Player];//正當性平均分數
                lb_Performance_Average.Text = comt.Form1.ImportData[7, comt.Form1.Now_Player];//表現性平均分數
                lb_CorrectnessAveragePlusPerformanceAverage.Text = ((double)(System.Convert.ToDouble(comt.Form1.ImportData[6, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[7, comt.Form1.Now_Player]))).ToString("0.00");//1st正當 + 1st表現
            }
            if (Select_Pose == "2" && Now_Pose == 1)
            {
                //6:1st正當性平均分數 
                //7:1st表現性平均分數 
                //8:第一品勢總分 
                //9:2nd正當性平均分數
                //10:2nd表現性平均分數
                //11:第二品勢總分 
                lb_Pose1_Totle.Text = comt.Form1.ImportData[8, comt.Form1.Now_Player];//第一品勢總分 = 1st正當性+1st表現性總分
                lb_Correctness_Average.Text = comt.Form1.ImportData[6, comt.Form1.Now_Player];//正當性平均分數
                lb_Performance_Average.Text = comt.Form1.ImportData[7, comt.Form1.Now_Player];//表現性平均分數
                lb_CorrectnessAveragePlusPerformanceAverage.Text = ((double)(System.Convert.ToDouble(comt.Form1.ImportData[6, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[7, comt.Form1.Now_Player]))).ToString("0.00");//1st正當 + 1st表現
            }
            if (Select_Pose == "2" && Now_Pose == 2)
            {
                lb_Pose1_Totle.Text = comt.Form1.ImportData[8, comt.Form1.Now_Player];//第一品勢總分
                lb_Pose2_Totle.Text = comt.Form1.ImportData[11, comt.Form1.Now_Player];//第二品勢總分
                
                lb_Correctness_Average.Text = ((double)(System.Convert.ToDouble(comt.Form1.ImportData[6, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[9, comt.Form1.Now_Player]))/2).ToString("0.00");
                lb_Performance_Average.Text = ((double)(System.Convert.ToDouble(comt.Form1.ImportData[7, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[10, comt.Form1.Now_Player]))/2).ToString("0.00");
                //lb_CorrectnessAveragePlusPerformanceAverage.Text = (((double)(System.Convert.ToDouble(comt.Form1.ImportData[6, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[9, comt.Form1.Now_Player])) / 2) + ((double)(System.Convert.ToDouble(comt.Form1.ImportData[7, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[10, comt.Form1.Now_Player])) / 2)).ToString("0.00");//原版
                lb_CorrectnessAveragePlusPerformanceAverage.Text = ((System.Convert.ToDouble(comt.Form1.ImportData[8, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[11, comt.Form1.Now_Player])) / 2).ToString("0.00");//與主控台一致版
               
                //原始分數//////////////////////////////////////////////////////
                lb_Original_Score_SumAll_Pose1.Text = comt.Form1.ImportData[19, comt.Form1.Now_Player];
                lb_Original_Score_SumAll_Pose2.Text = comt.Form1.ImportData[20, comt.Form1.Now_Player];
                lb_Original_Score_SumAll_Pose1AddPose2.Text = comt.Form1.ImportData[21, comt.Form1.Now_Player];
                //原始分數//////////////////////////////////////////////////////
            }

            if (Select_Pose == "2" && Now_Pose == 1)//這種情形不會開啟rank
            {

            }
            else 
            {
                timer_timer_ForCloseResult.Enabled = true;
            }

            //setFont////////////////////////////////////////////////////////////
            lb_Number.Font = new Font(lb_Number.Font.FontFamily, 100.0f, lb_Number.Font.Style);
            while (lb_Number.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Number.Text,
            new Font(lb_Number.Font.FontFamily, lb_Number.Font.Size, lb_Number.Font.Style)).Width)
            {
                lb_Number.Font = new Font(lb_Number.Font.FontFamily, lb_Number.Font.Size - 0.5f, lb_Number.Font.Style);
            }
        }
        public static Rank Rank;
        private void Result_FormClosing(object sender, FormClosingEventArgs e)
        {
            comt.Form1.IsOpen_Rank = true;//通知form1在下一個人開啟Demonstrate時將Rank關閉
            if (Select_Pose == "1" && Now_Pose == 1)//只有在 只比一個品勢 且一個品勢 已經比完 才會開啟結果視窗
            {
                Rank = new Rank();
                Rank.Show();
                //Thread.Sleep(50);                
            }
            if (Select_Pose == "2" && Now_Pose == 2)//只有在 比2個品勢 且2個品勢 都已經比完
            {
                Rank = new Rank();
                Rank.Show();
                //Thread.Sleep(50);                
            }
        }

        private void Result_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)//按下ESC  
            {
                this.Close();
            }
        }

        int Count_Second = 0;
        private void timer_timer_ForCloseResult_Tick(object sender, EventArgs e)
        {
            Count_Second++;
            if (Count_Second >= 3)
            {
                if (Select_Pose == "1" && Now_Pose == 1)//若只有一個品勢且已經比完則顯示玩resutl直接跳rank
                {
                    timer_timer_ForCloseResult.Enabled = false;
                    this.Close();
                }
                if (Select_Pose == "2" && Now_Pose == 1)//若有2品勢
                {
                    timer_timer_ForCloseResult.Enabled = false;
                }
                if (Select_Pose == "2" && Now_Pose == 2)
                {
                    timer_timer_ForCloseResult.Enabled = false;
                    this.Close();
                }
            }
        }
    }
}
