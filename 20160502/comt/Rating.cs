using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//-----------------------------------
using System.Threading;

namespace comt
{
    public partial class Rating : Form
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

        //翻牌專用///////////////////////////////////////////////////////////////        
        string[] Flag_Correctness;//看正確性評分了沒
        string[] Flag_Performance01;//看表現性性評分了沒
        string[] Flag_Performance02;//看表現性性評分了沒
        string[] Flag_Performance03;//看表現性性評分了沒
        float[] Score_Correctness;//填入正確性分數
        float[] Score_Performance01;//填入表現性分數
        float[] Score_Performance02;//填入表現性分數
        float[] Score_Performance03;//填入表現性分數
        //----------------------------------------------------------------------
        int count_Correctness_AlreadyRating = 0;//偵測已經評分的正當性有幾個，若已經評分個數=裁判個數，則代表評分完畢 
        int count_Performance01_AlreadyRating = 0;//偵測已經評分的表現性有幾個，若已經評分個數=裁判個數，則代表評分完畢 
        int count_Performance02_AlreadyRating = 0;//偵測已經評分的表現性有幾個，若已經評分個數=裁判個數，則代表評分完畢 
        int count_Performance03_AlreadyRating = 0;//偵測已經評分的表現性有幾個，若已經評分個數=裁判個數，則代表評分完畢 
        //翻牌專用///////////////////////////////////////////////////////////////
        Label[][] Score_Performance = new Label[3][];
        

        //狀態////////////////////////////////////////////////////////////////////////////////////////////////////
        //bool IsRating = false;
        //狀態////////////////////////////////////////////////////////////////////////////////////////////////////

        //找出最高值最低值，分數排序專用///////////////////////////////////////////////////////////////////////////////
        class ScoreRank_Correctness
        {
            public string Referee;
            public double Score;
            public ScoreRank_Correctness(string Referee, double Score)
            {
                this.Referee = Referee;
                this.Score = Score;
            }
        }
        ScoreRank_Correctness[] ScoreRank_Correctness_Variable;
        class ScoreRank_Performance
        {
            public string Referee;
            public double Score;
            public ScoreRank_Performance(string Referee, double Score)
            {
                this.Referee = Referee;
                this.Score = Score;
            }
        }
        ScoreRank_Performance[] ScoreRank_Performance_Variable;

        private int SortMethod_Correctness(ScoreRank_Correctness ScoreRank_Correctness1, ScoreRank_Correctness ScoreRank_Correctness2)
        {
            return ScoreRank_Correctness1.Score.CompareTo(ScoreRank_Correctness2.Score);
        }

        private int SortMethod_Performance(ScoreRank_Performance ScoreRank_Performance1, ScoreRank_Performance ScoreRank_Performance2)
        {
            return ScoreRank_Performance1.Score.CompareTo(ScoreRank_Performance2.Score);
        }

        string Highest_Score_Correctness;//最高分
        string Lowest_Score_Correctness;//最低分
        string Highest_Score_Performance;//最高分
        string Lowest_Score_Performance;//最低分
        string String_Score_Average_Correctness;//正確性平均分數
        string String_Score_Average_Performance;// = new string[0];//表現性平均分數     
        double Double_Score_Average_Correctness = 0;//正確性平均分數
        double[] Double_Score_Average_Performance;// = new double[0];//表現性平均分數  
        double Double_TotalScore_Correctness_Plus_Performance;//正確性+表現性的總分
        string String_TotalScore_Correctness_Plus_Performance;//正確性+表現性的總分        
        //找出最高值最低值，分數排序專用///////////////////////////////////////////////////////////////////////////////
        //新版原始總分
        double Original_Double_Score_SumAll_Correctness = 0;//正確性原始總分
        double[] Original_Double_Score_SumAll_Performance;//表現性原始總分 
        string Original_String_Score_SumAll_Pose1;//第一品勢原始總分
        string Original_String_Score_SumAll_Pose2;//第二品勢原始總分
        string Original_String_Score_SumAll_Pose1AddPose2;//第一品勢+第二品勢原始總分        
        //新版原始總分

        //字型顏色//////////////////////////////////////////////////////
        Color Show_Before = Color.CadetBlue;//Color.Gray;
        Color Show_After = Color.Yellow;//Color.White;
        Color Show_Ok = Color.Pink;//Color.White;
        //字型顏色//////////////////////////////////////////////////////


        public Rating()
        {    
            //移動視窗到延伸桌面
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            InitializeComponent();
            timer_Rating.Start();
            Score_Performance[0] = new Label[7] { lb_Performance01_Referee01, lb_Performance01_Referee02, lb_Performance01_Referee03, lb_Performance01_Referee04, lb_Performance01_Referee05, lb_Performance01_Referee06, lb_Performance01_Referee07 };
            Score_Performance[1] = new Label[7] { lb_Performance02_Referee01, lb_Performance02_Referee02, lb_Performance02_Referee03, lb_Performance02_Referee04, lb_Performance02_Referee05, lb_Performance02_Referee06, lb_Performance02_Referee07 };
            Score_Performance[2] = new Label[7] { lb_Performance03_Referee01, lb_Performance03_Referee02, lb_Performance03_Referee03, lb_Performance03_Referee04, lb_Performance03_Referee05, lb_Performance03_Referee06, lb_Performance03_Referee07 };
            //ColorIni///////////////////////////////////////////////////////////////////////
            lb_Correctness_Referee01.ForeColor = Show_Before;
            lb_Correctness_Referee02.ForeColor = Show_Before;
            lb_Correctness_Referee03.ForeColor = Show_Before;
            lb_Correctness_Referee04.ForeColor = Show_Before;
            lb_Correctness_Referee05.ForeColor = Show_Before;
            lb_Correctness_Referee06.ForeColor = Show_Before;
            lb_Correctness_Referee07.ForeColor = Show_Before;
            //----------
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    Score_Performance[i][j].ForeColor = Show_Before;
                }
            }

            
            Double_Score_Average_Performance = new double[3] { 0, 0, 0 };
            Original_Double_Score_SumAll_Performance = new double[3] { 0, 0, 0 };
            //----------
            lb_Pose_TotalScore.ForeColor = Show_Before;

            Count_Referee = comt.Form1.Count_Referee;//有幾位評審
            //
            
            if (Count_Referee <= 5)
            {
                //TempletIni///////////////////////////////////////////////////////////////////////////
                #region 排版
                //背景圖片///////////////////////////////////////////////////////
                this.BackgroundImage = global::comt.Properties.Resources.Rating4Columns_M5;
                //背景圖片///////////////////////////////////////////////////////

                //正確性字////////////////////////////////////////////////
                this.lb_Correctness_Referee01.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee01.Location = new System.Drawing.Point(80, 39);
                this.lb_Correctness_Referee01.Size = new System.Drawing.Size(159, 101);

                this.lb_Correctness_Referee02.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee02.Location = new System.Drawing.Point(80, 183);
                this.lb_Correctness_Referee02.Size = new System.Drawing.Size(159, 101);

                this.lb_Correctness_Referee03.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee03.Location = new System.Drawing.Point(80, 334);
                this.lb_Correctness_Referee03.Size = new System.Drawing.Size(159, 101);

                this.lb_Correctness_Referee04.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee04.Location = new System.Drawing.Point(80, 477);
                this.lb_Correctness_Referee04.Size = new System.Drawing.Size(159, 101);

                this.lb_Correctness_Referee05.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee05.Location = new System.Drawing.Point(80, 627);
                this.lb_Correctness_Referee05.Size = new System.Drawing.Size(159, 101);

                this.lb_Correctness_Referee06.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee06.Location = new System.Drawing.Point(1089, 215);
                this.lb_Correctness_Referee06.Size = new System.Drawing.Size(205, 130);

                this.lb_Correctness_Referee07.Font = new System.Drawing.Font("微軟正黑體", 78F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee07.Location = new System.Drawing.Point(1089, 320);
                this.lb_Correctness_Referee07.Size = new System.Drawing.Size(205, 130);
                //正確性字////////////////////////////////////////////////

                //表現性1字////////////////////////////////////////////////
                this.lb_Performance01_Referee01.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee01.Location = new System.Drawing.Point(213, 39);
                this.lb_Performance01_Referee01.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance01_Referee02.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee02.Location = new System.Drawing.Point(213, 183);
                this.lb_Performance01_Referee02.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance01_Referee03.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee03.Location = new System.Drawing.Point(213, 334);
                this.lb_Performance01_Referee03.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance01_Referee04.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee04.Location = new System.Drawing.Point(213, 477);
                this.lb_Performance01_Referee04.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance01_Referee05.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee05.Location = new System.Drawing.Point(213, 627);
                this.lb_Performance01_Referee05.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance01_Referee06.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee06.Location = new System.Drawing.Point(1271, 217);
                this.lb_Performance01_Referee06.Size = new System.Drawing.Size(205, 130);
                this.lb_Performance01_Referee06.Visible = false;

                this.lb_Performance01_Referee07.Font = new System.Drawing.Font("微軟正黑體", 78F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance01_Referee07.Location = new System.Drawing.Point(1271, 322);
                this.lb_Performance01_Referee07.Size = new System.Drawing.Size(205, 130);
                this.lb_Performance01_Referee07.Visible = false;
                //表現性1字////////////////////////////////////////////////


                //表現性2字////////////////////////////////////////////////
                this.lb_Performance02_Referee01.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee01.Location = new System.Drawing.Point(346, 39);
                this.lb_Performance02_Referee01.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance02_Referee02.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee02.Location = new System.Drawing.Point(346, 183);
                this.lb_Performance02_Referee02.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance02_Referee03.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee03.Location = new System.Drawing.Point(346, 334);
                this.lb_Performance02_Referee03.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance02_Referee04.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee04.Location = new System.Drawing.Point(346, 477);
                this.lb_Performance02_Referee04.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance02_Referee05.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee05.Location = new System.Drawing.Point(346, 627);
                this.lb_Performance02_Referee05.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance02_Referee06.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee06.Location = new System.Drawing.Point(1271, 217);
                this.lb_Performance02_Referee06.Size = new System.Drawing.Size(205, 130);
                this.lb_Performance02_Referee06.Visible = false;

                this.lb_Performance02_Referee07.Font = new System.Drawing.Font("微軟正黑體", 78F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance02_Referee07.Location = new System.Drawing.Point(1271, 322);
                this.lb_Performance02_Referee07.Size = new System.Drawing.Size(205, 130);
                this.lb_Performance02_Referee07.Visible = false;
                //表現性2字////////////////////////////////////////////////


                //表現性3字////////////////////////////////////////////////
                this.lb_Performance03_Referee01.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee01.Location = new System.Drawing.Point(481, 39);
                this.lb_Performance03_Referee01.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance03_Referee02.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee02.Location = new System.Drawing.Point(481, 183);
                this.lb_Performance03_Referee02.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance03_Referee03.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee03.Location = new System.Drawing.Point(481, 334);
                this.lb_Performance03_Referee03.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance03_Referee04.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee04.Location = new System.Drawing.Point(481, 477);
                this.lb_Performance03_Referee04.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance03_Referee05.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee05.Location = new System.Drawing.Point(481, 627);
                this.lb_Performance03_Referee05.Size = new System.Drawing.Size(159, 101);

                this.lb_Performance03_Referee06.Font = new System.Drawing.Font("微軟正黑體", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee06.Location = new System.Drawing.Point(1271, 217);
                this.lb_Performance03_Referee06.Size = new System.Drawing.Size(205, 130);
                this.lb_Performance03_Referee06.Visible = false;

                this.lb_Performance03_Referee07.Font = new System.Drawing.Font("微軟正黑體", 78F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance03_Referee07.Location = new System.Drawing.Point(1271, 322);
                this.lb_Performance03_Referee07.Size = new System.Drawing.Size(205, 130);
                this.lb_Performance03_Referee07.Visible = false;
                //表現性3字////////////////////////////////////////////////
                #region mark
            
                #endregion
                #endregion
                //TempletIni///////////////////////////////////////////////////////////////////////////
            }
            


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

            //初始化版面看有幾個評審就用幾個顯示，沒用到的就空白///////////////////////////////////////
            #region 初始化 正當性
            if (1 > Count_Referee)
            {
                lb_Correctness_Referee01.Text = "";
            }
            if (2 > Count_Referee)
            {
                lb_Correctness_Referee02.Text = "";
            }
            if (3 > Count_Referee)
            {
                lb_Correctness_Referee03.Text = "";
            }
            if (4 > Count_Referee)
            {
                lb_Correctness_Referee04.Text = "";
            }
            if (5 > Count_Referee)
            {
                lb_Correctness_Referee05.Text = "";
            }
            if (6 > Count_Referee)
            {
                lb_Correctness_Referee06.Text = "";
            }
            if (7 > Count_Referee)
            {
                lb_Correctness_Referee07.Text = "";
            }
            #endregion
            #region 初始化 表現性
            if (1 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][0].Text = "";
                }
                
            }
            if (2 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][1].Text = "";
                }
            }
            if (3 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][2].Text = "";
                }
            }
            if (4 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][3].Text = "";
                }
            }
            if (5 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][4].Text = "";
                }
            }
            if (6 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][5].Text = "";
                }
            }
            if (7 > Count_Referee)
            {
                for (int i = 0; i < 3; i++)
                {
                    Score_Performance[i][6].Text = "";
                }
            }
            #endregion
            #region 初始化 評審編號
            if (1 > Count_Referee)
            {
                lb_Referee01_Title.Text = "";
            }
            if (2 > Count_Referee)
            {
                lb_Referee02_Title.Text = "";
            }
            if (3 > Count_Referee)
            {
                lb_Referee03_Title.Text = "";
            }
            if (4 > Count_Referee)
            {
                lb_Referee04_Title.Text = "";
            }
            if (5 > Count_Referee)
            {
                lb_Referee05_Title.Text = "";
            }
            if (6 > Count_Referee)
            {
                lb_Referee06_Title.Text = "";
            }
            if (7 > Count_Referee)
            {
                lb_Referee07_Title.Text = "";
            }
            #endregion
            //初始化版面看有幾個評審就用幾個顯示，沒用到的就空白///////////////////////////////////////

            //排名刪除最高最低專用/////////////////////////////////////////////////////////////////////
            ScoreRank_Correctness_Variable = new ScoreRank_Correctness[Count_Referee + 1];//多一個位置是用來放-1，放在[0]因為-1一定是最小的分數
            ScoreRank_Performance_Variable = new ScoreRank_Performance[Count_Referee + 1];//多一個位置是用來放-1，放在[0]因為-1一定是最小的分數
            //排名刪除最高最低專用/////////////////////////////////////////////////////////////////////

            if (comt.Form1.Now_Pose.ToString() == "1")
            {                
                lb_ID.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名  
                //lb_Pose.Text = "第一品勢";//lb_NowPose
                //lb_NowPose.Text = "第一品勢";//lb_NowPose
                lb_NowPose.Text = "POOMSAE 1";//lb_NowPose
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
            }
            if (comt.Form1.Now_Pose.ToString() == "2")
            {                
                lb_ID.Text = comt.Form1.ImportData[15, comt.Form1.Now_Player];//編號
                lb_Group.Text = comt.Form1.ImportData[0, comt.Form1.Now_Player];//組別
                lb_Unit.Text = comt.Form1.ImportData[1, comt.Form1.Now_Player];//單位
                lb_Name.Text = comt.Form1.ImportData[2, comt.Form1.Now_Player];//姓名
                //lb_Pose.Text = "第二品勢";
                //lb_NowPose.Text = "第二品勢";//lb_NowPose
                lb_NowPose.Text = "POOMSAE 2";//lb_NowPose
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
            }
        }

        private void timer_Rating_Tick(object sender, EventArgs e)
        {
            Flag_Correctness = comt.Form1.Flag_Correctness;
            Flag_Performance01 = comt.Form1.Flag_Performance01;
            Flag_Performance02 = comt.Form1.Flag_Performance02;
            Flag_Performance03 = comt.Form1.Flag_Performance03;
            Score_Correctness = comt.Form1.ScoreArray[0];
            Score_Performance01 = comt.Form1.ScoreArray[1];
            Score_Performance02 = comt.Form1.ScoreArray[2];
            Score_Performance03 = comt.Form1.ScoreArray[3];
            count_Correctness_AlreadyRating = 0;//每次進來要歸零才可以讓下面的for做該次的已評分個數統計
            for (int i = 1; i <= Count_Referee; i++)
            {
                if (Flag_Correctness[i] == "1")//去統計已經有幾位評審評分了
                {
                    count_Correctness_AlreadyRating++;
                }
            }

            if (count_Correctness_AlreadyRating == Count_Referee)//若已經評分個數=裁判個數，則代表評分完畢
            {//那就翻牌            

                //將所有數值填入class的陣列內，準備做排序
                for (int i = 0; i <= Count_Referee; i++)
                {
                    if (i == 0)
                    {
                        ScoreRank_Correctness_Variable[0] = new ScoreRank_Correctness("0", -1);//因為0的位置一定要放東西才能Sort，所以放一個-1給他，因為分數一定不會是負數，所以之後sort出來的陣列的位置[0]一定是最小值，也就是-1，所以我們要的真正最小值在位置[1]
                    }
                    else
                    {
                        ScoreRank_Correctness_Variable[i] = new ScoreRank_Correctness(i.ToString(), System.Convert.ToDouble(Score_Correctness[i]));
                    }
                }

                Array.Sort(ScoreRank_Correctness_Variable, SortMethod_Correctness);
                Lowest_Score_Correctness = ScoreRank_Correctness_Variable[1].Referee;//排序完以後[0]的位置一定是-1，[1]的位置一定是最小值
                Highest_Score_Correctness = ScoreRank_Correctness_Variable[Count_Referee].Referee;
                Double_Score_Average_Correctness = 0;//初始化
                //算正當性分數平均/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (Count_Referee == 1)//如果只有一個人,最低分就是平均分數，ie,sort完以後的位置[1]
                {
                    Double_Score_Average_Correctness = ScoreRank_Correctness_Variable[1].Score;
                    String_Score_Average_Correctness = ScoreRank_Correctness_Variable[1].Score.ToString("0.00");//String_Score_Average_Correctness = Lowest_Score_Correctness;
                }
                else if (Count_Referee == 2)//如果有兩個人，就( 位置[1] +位置[2] )在除以 2
                {
                    Double_Score_Average_Correctness = ((double)((ScoreRank_Correctness_Variable[1].Score + ScoreRank_Correctness_Variable[2].Score) / 2));
                    String_Score_Average_Correctness = ((double)((ScoreRank_Correctness_Variable[1].Score + ScoreRank_Correctness_Variable[2].Score) / 2)).ToString("0.00");
                }
                else if (Count_Referee == 3)//如果有三個人，就從位置[2] 加到 位置[最高位-1] 因為要 去掉"最高" 和 "最低"(ie,位置0)
                {
                    Double_Score_Average_Correctness = ((double)((ScoreRank_Correctness_Variable[1].Score + ScoreRank_Correctness_Variable[2].Score + ScoreRank_Correctness_Variable[3].Score) / 3));
                    String_Score_Average_Correctness = ((double)((ScoreRank_Correctness_Variable[1].Score + ScoreRank_Correctness_Variable[2].Score + ScoreRank_Correctness_Variable[3].Score) / 3)).ToString("0.00");
                }
                else if (Count_Referee > 3)
                {
                    for (int i = 2; i <= (Count_Referee-1); i++)
                    {
                        Double_Score_Average_Correctness = Double_Score_Average_Correctness + ScoreRank_Correctness_Variable[i].Score;
                    }
                    Double_Score_Average_Correctness = Double_Score_Average_Correctness / (Count_Referee - 2);
                    String_Score_Average_Correctness = Double_Score_Average_Correctness.ToString("0.00");
                }
                lb_Correctness_Score.ForeColor = Show_After;
                lb_Correctness_Score.Text = String_Score_Average_Correctness;
                //算正當性分數平均/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                if (1 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)//看有幾位評審，就翻幾個牌，因為翻1號評審，所以拿1去跟評審個數比，若不大於就翻牌
                {
                    lb_Correctness_Referee01.ForeColor = Show_After;
                    lb_Correctness_Referee01.Text = Score_Correctness[1].ToString("0.0");
                    if (("1" == Lowest_Score_Correctness || "1" == Highest_Score_Correctness) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Correctness_Referee01.ForeColor = Color.Red;
                    }
                }
                if (2 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee02.ForeColor = Show_After;
                    lb_Correctness_Referee02.Text = Score_Correctness[2].ToString("0.0");
                    if (("2" == Lowest_Score_Correctness || "2" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee02.ForeColor = Color.Red;
                    }
                }
                if (3 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee03.ForeColor = Show_After;
                    lb_Correctness_Referee03.Text = Score_Correctness[3].ToString("0.0");
                    if (("3" == Lowest_Score_Correctness || "3" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee03.ForeColor = Color.Red;
                    }
                }
                if (4 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee04.ForeColor = Show_After;
                    lb_Correctness_Referee04.Text = Score_Correctness[4].ToString("0.0");
                    if (("4" == Lowest_Score_Correctness || "4" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee04.ForeColor = Color.Red;
                    }
                }
                if (5 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee05.ForeColor = Show_After;
                    lb_Correctness_Referee05.Text = Score_Correctness[5].ToString("0.0");
                    if (("5" == Lowest_Score_Correctness || "5" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee05.ForeColor = Color.Red;
                    }
                }
                if (6 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee06.ForeColor = Show_After;
                    lb_Correctness_Referee06.Text = Score_Correctness[6].ToString("0.0");
                    if (("6" == Lowest_Score_Correctness || "6" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee06.ForeColor = Color.Red;
                    }
                }
                if (7 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee07.ForeColor = Show_After;
                    lb_Correctness_Referee07.Text = Score_Correctness[7].ToString("0.0");
                    if (("7" == Lowest_Score_Correctness || "7" == Highest_Score_Correctness) && Count_Referee >= 3)
                    {
                        lb_Correctness_Referee07.ForeColor = Color.Red;
                    }
                }
            }
            else
            {                
                if (Flag_Correctness[1] == "1" && 1 <= Count_Referee)
                {
                    lb_Correctness_Referee01.ForeColor = Show_Ok;
                    lb_Correctness_Referee01.Text = "OK";
                }
                if (Flag_Correctness[2] == "1" && 2 <= Count_Referee)
                {
                    lb_Correctness_Referee02.ForeColor = Show_Ok;
                    lb_Correctness_Referee02.Text = "OK";
                }
                if (Flag_Correctness[3] == "1" && 3 <= Count_Referee)
                {
                    lb_Correctness_Referee03.ForeColor = Show_Ok;
                    lb_Correctness_Referee03.Text = "OK";
                }
                if (Flag_Correctness[4] == "1" && 4 <= Count_Referee)
                {
                    lb_Correctness_Referee04.ForeColor = Show_Ok;
                    lb_Correctness_Referee04.Text = "OK";
                }
                if (Flag_Correctness[5] == "1" && 5 <= Count_Referee)
                {
                    lb_Correctness_Referee05.ForeColor = Show_Ok;
                    lb_Correctness_Referee05.Text = "OK";
                }
                if (Flag_Correctness[6] == "1" && 6 <= Count_Referee)
                {
                    lb_Correctness_Referee06.ForeColor = Show_Ok;
                    lb_Correctness_Referee06.Text = "OK";
                }
                if (Flag_Correctness[7] == "1" && 7 <= Count_Referee)
                {
                    lb_Correctness_Referee07.ForeColor = Show_Ok;
                    lb_Correctness_Referee07.Text = "OK";
                }
            }
            scoring_performance(count_Performance01_AlreadyRating, Flag_Performance01, Score_Performance01, Score_Performance[0], 0);
            scoring_performance(count_Performance02_AlreadyRating, Flag_Performance02, Score_Performance02, Score_Performance[1], 1);
            scoring_performance(count_Performance03_AlreadyRating, Flag_Performance03, Score_Performance03, Score_Performance[2], 2);
        }

        private void scoring_performance(int count_Performance_AlreadyRating, string[] Flag_Performance, float[] Score_Performance, System.Windows.Forms.Label[] lb_Performance_Referee, int Scoring_PerformanceNum)
        {
            count_Performance_AlreadyRating = 0;//每次進來要歸零才可以讓下面的for做該次的已評分個數統計
            for (int i = 1; i <= Count_Referee; i++)
            {
                if (Flag_Performance[i] == "1")//去統計已經有幾位評審評分了，若已經評分個數=裁判個數，則代表評分完畢
                {
                    count_Performance_AlreadyRating++;
                }
            }
         
            if (count_Performance_AlreadyRating == Count_Referee)//不須正確性全完才翻牌，表現性以評分可先翻//原本if (count_Performance_AlreadyRating == Count_Referee && count_Correctness_AlreadyRating == Count_Referee)//若已經評分個數=裁判個數，則代表評分完畢//正確顯示完再顯示表現
            {//那就翻牌                
                //將所有數值填入class的陣列內，準備做排序
                for (int i = 0; i <= Count_Referee; i++)
                {
                    if (i == 0)
                    {
                        ScoreRank_Performance_Variable[0] = new ScoreRank_Performance("0", -1);//因為0的位置一定要放東西才能Sort，所以放一個-1給他，因為分數一定不會是負數，所以之後sort出來的陣列的位置[0]一定是最小值，也就是-1，所以我們要的真正最小值在位置[1]
                    }
                    else
                    {
                        ScoreRank_Performance_Variable[i] = new ScoreRank_Performance(i.ToString(), System.Convert.ToDouble(Score_Performance[i]));
                    }
                }

                Array.Sort(ScoreRank_Performance_Variable, SortMethod_Performance);
                Lowest_Score_Performance = ScoreRank_Performance_Variable[1].Referee;//排序完以後[0]的位置一定是-1，[1]的位置一定是最小值
                Highest_Score_Performance = ScoreRank_Performance_Variable[Count_Referee].Referee;



                Double_Score_Average_Performance[Scoring_PerformanceNum] = 0;//初始化
                //算表現性分數平均/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (Count_Referee == 1)//如果只有一個人,最低分就是平均分數，ie,sort完以後的位置[1]
                {
                    Double_Score_Average_Performance[Scoring_PerformanceNum] = ScoreRank_Performance_Variable[1].Score;
                    String_Score_Average_Performance = Double_Score_Average_Performance.Sum().ToString("0.00");//String_Score_Average_Performance = Lowest_Score_Performance;
                }
                else if (Count_Referee == 2)//如果有兩個人，就( 位置[1] +位置[2] )在除以 2
                {
                    Double_Score_Average_Performance[Scoring_PerformanceNum] = ((double)((ScoreRank_Performance_Variable[1].Score + ScoreRank_Performance_Variable[2].Score) / 2));
                    String_Score_Average_Performance = Double_Score_Average_Performance.Sum().ToString("0.00");//.ToString("0.00")代表四捨五入到小數點後面第2位
                }
                else if (Count_Referee == 3)
                {
                    Double_Score_Average_Performance[Scoring_PerformanceNum] = ((double)((ScoreRank_Performance_Variable[1].Score + ScoreRank_Performance_Variable[2].Score + ScoreRank_Performance_Variable[3].Score) / 3));
                    String_Score_Average_Performance = Double_Score_Average_Performance.Sum().ToString("0.00");
                }
                else if (Count_Referee > 3)//如果有三個人以上，就從位置[2] 加到 位置[最高位-1] 因為要 去掉"最高" 和 "最低"(ie,位置0)
                {
                    for (int i = 2; i <= (Count_Referee - 1); i++)
                    {
                        Double_Score_Average_Performance[Scoring_PerformanceNum] = Double_Score_Average_Performance[Scoring_PerformanceNum] + ScoreRank_Performance_Variable[i].Score;
                    }
                    Double_Score_Average_Performance[Scoring_PerformanceNum] = Double_Score_Average_Performance[Scoring_PerformanceNum] / (Count_Referee - 2);
                    String_Score_Average_Performance = Double_Score_Average_Performance.Sum().ToString("0.00");
                }
                //lb_Performance_Score.ForeColor = Show_After;
                
                //算表現性分數平均/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Original_Double_Score_SumAll_Correctness = 0;
                Original_Double_Score_SumAll_Performance[Scoring_PerformanceNum] = 0;
                //計算原始總分////////////////////////////////////
                for (int i = 1; i <= Count_Referee; i++)
                {
                    Original_Double_Score_SumAll_Correctness = Original_Double_Score_SumAll_Correctness + ScoreRank_Correctness_Variable[i].Score;
                    Original_Double_Score_SumAll_Performance[Scoring_PerformanceNum] = Original_Double_Score_SumAll_Performance[Scoring_PerformanceNum] + ScoreRank_Performance_Variable[i].Score;
                }
                //計算原始總分////////////////////////////////////

                Double_TotalScore_Correctness_Plus_Performance = Double_Score_Average_Correctness + Double_Score_Average_Performance.Sum();
                String_TotalScore_Correctness_Plus_Performance = Double_TotalScore_Correctness_Plus_Performance.ToString("0.00");
                if (count_Performance_AlreadyRating == Count_Referee && Scoring_PerformanceNum == 2)
                {
                    lb_Performance_Score.ForeColor = Show_After;
                    lb_Performance_Score.Text = String_Score_Average_Performance;
                    //全部分數Data只在這裡進行寫入///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    if (Now_Pose == 1)
                    {
                        comt.Form1.IsWritng_Rank = true;//正在寫入
                        comt.Form1.ImportData[6, comt.Form1.Now_Player] = String_Score_Average_Correctness;//正當性平均分數                    
                        comt.Form1.ImportData[7, comt.Form1.Now_Player] = String_Score_Average_Performance;//表現性平均分數
                        comt.Form1.ImportData[8, comt.Form1.Now_Player] = String_TotalScore_Correctness_Plus_Performance;//正當性+表現性總分
                        lb_Pose_TotalScore.ForeColor = Show_After;
                        if (Double_TotalScore_Correctness_Plus_Performance >= 10)
                        {
                            String_TotalScore_Correctness_Plus_Performance = Double_TotalScore_Correctness_Plus_Performance.ToString("0.0");
                        }
                        lb_Pose_TotalScore.Text = String_TotalScore_Correctness_Plus_Performance;
                        if (Scoring_PerformanceNum == 2)  // 評完所有表現性
                            timer_ForCloseWinForm.Enabled = true;
                        //新加入第一品勢原始分數欄位////////////////////////////////////////////////////////////////
                        Original_String_Score_SumAll_Pose1 = (Original_Double_Score_SumAll_Correctness + Original_Double_Score_SumAll_Performance.Sum()).ToString("0.00");
                        lb_Original_Score_SumAll.Text = Original_String_Score_SumAll_Pose1;
                        comt.Form1.ImportData[19, comt.Form1.Now_Player] = Original_String_Score_SumAll_Pose1;//第一品勢原始總分

                    }
                    if (Now_Pose == 2)
                    {
                        comt.Form1.ImportData[9, comt.Form1.Now_Player] = String_Score_Average_Correctness;//正當性平均分數
                        comt.Form1.ImportData[10, comt.Form1.Now_Player] = String_Score_Average_Performance;//表現性平均分數
                        comt.Form1.ImportData[11, comt.Form1.Now_Player] = String_TotalScore_Correctness_Plus_Performance;//正當性+表現性總分
                        lb_Pose_TotalScore.ForeColor = Show_After;
                        if (Double_TotalScore_Correctness_Plus_Performance >= 10)
                        {
                            String_TotalScore_Correctness_Plus_Performance = Double_TotalScore_Correctness_Plus_Performance.ToString("0.0");
                        }
                        lb_Pose_TotalScore.Text = String_TotalScore_Correctness_Plus_Performance;
                        if (Scoring_PerformanceNum == 2)  // 評完所有表現性
                            timer_ForCloseWinForm.Enabled = true;
                        //新加入第二品勢原始分數欄位////////////////////////////////////////////////////////////////
                        Original_String_Score_SumAll_Pose2 = (Original_Double_Score_SumAll_Correctness + Original_Double_Score_SumAll_Performance.Sum()).ToString("0.00");
                        lb_Original_Score_SumAll.Text = Original_String_Score_SumAll_Pose2;
                        comt.Form1.ImportData[20, comt.Form1.Now_Player] = Original_String_Score_SumAll_Pose2;//第二品勢原始總分

                    }
                }
                    
                



                //全部分數Data只在這裡進行寫入///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (1 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)//看有幾位評審，就翻幾個牌，因為翻1號評審，所以拿1去跟評審個數比，若不大於就翻牌
                {
                    lb_Performance_Referee[0].ForeColor = Show_After;
                    lb_Performance_Referee[0].Text = Score_Performance[1].ToString("0.0");
                    if (("1" == Lowest_Score_Performance || "1" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[0].ForeColor = Color.Red;
                    }
                }
                if (2 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee[1].ForeColor = Show_After;
                    lb_Performance_Referee[1].Text = Score_Performance[2].ToString("0.0");
                    if (("2" == Lowest_Score_Performance || "2" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[1].ForeColor = Color.Red;
                    }
                }
                if (3 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee[2].ForeColor = Show_After;
                    lb_Performance_Referee[2].Text = Score_Performance[3].ToString("0.0");
                    if (("3" == Lowest_Score_Performance || "3" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[2].ForeColor = Color.Red;
                    }
                }
                if (4 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee[3].ForeColor = Show_After;
                    lb_Performance_Referee[3].Text = Score_Performance[4].ToString("0.0");
                    if (("4" == Lowest_Score_Performance || "4" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[3].ForeColor = Color.Red;
                    }
                }
                if (5 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee[4].ForeColor = Show_After;
                    lb_Performance_Referee[4].Text = Score_Performance[5].ToString("0.0");
                    if (("5" == Lowest_Score_Performance || "5" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[4].ForeColor = Color.Red;
                    }
                }
                if (6 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee[5].ForeColor = Show_After;
                    lb_Performance_Referee[5].Text = Score_Performance[6].ToString("0.0");
                    if (("6" == Lowest_Score_Performance || "6" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[5].ForeColor = Color.Red;
                    }
                }
                if (7 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee[6].ForeColor = Show_After;
                    lb_Performance_Referee[6].Text = Score_Performance[7].ToString("0.0");
                    if (("7" == Lowest_Score_Performance || "7" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee[6].ForeColor = Color.Red;
                    }
                }
            }
            else
            {
                if (Flag_Performance[1] == "1" && 1 <= Count_Referee)
                {
                    lb_Performance_Referee[0].ForeColor = Show_Ok;
                    lb_Performance_Referee[0].Text = "OK";
                }
                if (Flag_Performance[2] == "1" && 2 <= Count_Referee)
                {
                    lb_Performance_Referee[1].ForeColor = Show_Ok;
                    lb_Performance_Referee[1].Text = "OK";
                }
                if (Flag_Performance[3] == "1" && 3 <= Count_Referee)
                {
                    lb_Performance_Referee[2].ForeColor = Show_Ok;
                    lb_Performance_Referee[2].Text = "OK";
                }
                if (Flag_Performance[4] == "1" && 4 <= Count_Referee)
                {
                    lb_Performance_Referee[3].ForeColor = Show_Ok;
                    lb_Performance_Referee[3].Text = "OK";
                }
                if (Flag_Performance[5] == "1" && 5 <= Count_Referee)
                {
                    lb_Performance_Referee[4].ForeColor = Show_Ok;
                    lb_Performance_Referee[4].Text = "OK";
                }
                if (Flag_Performance[6] == "1" && 6 <= Count_Referee)
                {
                    lb_Performance_Referee[5].ForeColor = Show_Ok;
                    lb_Performance_Referee[5].Text = "OK";
                }
                if (Flag_Performance[7] == "1" && 7 <= Count_Referee)
                {
                    lb_Performance_Referee[6].ForeColor = Show_Ok;
                    lb_Performance_Referee[6].Text = "OK";
                }
            }
        }
        public static Result Result;
        public static Rank Rank;
        private void Rating_FormClosing(object sender, FormClosingEventArgs e)
        { 
            if (Select_Pose == "1" && Now_Pose == 1)//只有在 只比一個品勢 且一個品勢 已經比完 才會開啟結果視窗
            {
                string temp_sdfg = comt.Form1.ImportData[8, comt.Form1.Now_Player];
                comt.Form1.ImportData[12, comt.Form1.Now_Player] = comt.Form1.ImportData[8, comt.Form1.Now_Player];//因為單一品勢，所以 總平均 = 單一品勢分數

                //新加入總原始分數欄位////////////////////////////////////////////////////////////////
                comt.Form1.ImportData[21, comt.Form1.Now_Player] = comt.Form1.ImportData[19, comt.Form1.Now_Player];//因為單一品勢，所以 原始總分 = 第一品勢的原始總分
                //新加入總原始分數欄位////////////////////////////////////////////////////////////////

                comt.Form1.IsWritng_Rank = true;//正在寫入
                Rank = new Rank();
                Rank.Show();
                comt.Form1.WriteExcelAll();
                comt.Form1.IsOpen_Rank_ByRating = true;
                
                //Thread.Sleep(2000);                
            }
            if (Select_Pose == "2" && Now_Pose == 1)
            {
                comt.Form1.WriteExcelAll();
                /*
                Result = new Result();
                Result.Show();
                */
                //Thread.Sleep(50);                
            }
            if (Select_Pose == "2" && Now_Pose == 2)//只有在 比2個品勢 且2個品勢 都已經比完
            {
                //開啟結果前將結果寫入總分
                string temp1 = comt.Form1.ImportData[8, comt.Form1.Now_Player];
                string temp2 = comt.Form1.ImportData[11, comt.Form1.Now_Player];
                comt.Form1.IsWritng_Rank = true;//正在寫入
                comt.Form1.ImportData[12, comt.Form1.Now_Player] = ((System.Convert.ToDouble(comt.Form1.ImportData[8, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[11, comt.Form1.Now_Player])) / 2).ToString("0.00");

                //新加入總原始分數欄位////////////////////////////////////////////////////////////////
                comt.Form1.ImportData[21, comt.Form1.Now_Player] = 
                (
                    System.Convert.ToDouble(comt.Form1.ImportData[19, comt.Form1.Now_Player]) 
                    +
                    System.Convert.ToDouble(comt.Form1.ImportData[20, comt.Form1.Now_Player])
                ).ToString("0.00");//因為雙品勢，所以 原始總分 = 第一品勢的原始總分 + 第二品勢的原始總分
                //新加入總原始分數欄位////////////////////////////////////////////////////////////////

                Result = new Result();
                Result.Show();
                //Thread.Sleep(50);                
            }
        }

        private void Rating_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)//按下ESC  
            {
                this.Close();
            }
        }

        int Count_Second = 0;
        private void timer_ForCloseWinForm_Tick(object sender, EventArgs e)
        {
            Count_Second++;//數秒，用以計算3秒鐘關閉視窗
            //Rank = new Rank();
            //Rank.Activate();
            if (Count_Second >= 3)
            {
                if (Select_Pose == "1" && Now_Pose == 1)
                {
                    timer_ForCloseWinForm.Enabled = false;
                    this.Close();
                }
                if (Select_Pose == "2" && Now_Pose == 1)
                {
                    timer_ForCloseWinForm.Enabled = false;
                }
                if (Select_Pose == "2" && Now_Pose == 2)
                {
                    timer_ForCloseWinForm.Enabled = false;
                    this.Close();
                }
            }
        }

    }
}

