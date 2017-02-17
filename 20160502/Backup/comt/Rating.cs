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
        string[] Flag_Performance;//看表現性性評分了沒
        string[] Score_Correctness;//填入正確性分數
        string[] Score_Performance;//填入表現性分數
        //----------------------------------------------------------------------
        int count_Correctness_AlreadyRating = 0;//偵測已經評分的正當性有幾個，若已經評分個數=裁判個數，則代表評分完畢 
        int count_Performance_AlreadyRating = 0;//偵測已經評分的表現性有幾個，若已經評分個數=裁判個數，則代表評分完畢 
        //翻牌專用///////////////////////////////////////////////////////////////

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
        string String_Score_Average_Performance;//表現性平均分數     
        double Double_Score_Average_Correctness = 0;//正確性平均分數
        double Double_Score_Average_Performance = 0;//表現性平均分數  
        double Double_TotalScore_Correctness_Plus_Performance;//正確性+表現性的總分
        string String_TotalScore_Correctness_Plus_Performance;//正確性+表現性的總分
        //找出最高值最低值，分數排序專用///////////////////////////////////////////////////////////////////////////////


        //字型顏色//////////////////////////////////////////////////////
        Color Show_Before = Color.Gray;
        Color Show_After = Color.White;
        //字型顏色//////////////////////////////////////////////////////


        public Rating()
        {
            InitializeComponent();
            timer_Rating.Enabled = true;

            //ColorIni///////////////////////////////////////////////////////////////////////
            lb_Correctness_Referee01.ForeColor = Show_Before;
            lb_Correctness_Referee02.ForeColor = Show_Before;
            lb_Correctness_Referee03.ForeColor = Show_Before;
            lb_Correctness_Referee04.ForeColor = Show_Before;
            lb_Correctness_Referee05.ForeColor = Show_Before;
            lb_Correctness_Referee06.ForeColor = Show_Before;
            lb_Correctness_Referee07.ForeColor = Show_Before;
            //----------
            lb_Performance_Referee01.ForeColor = Show_Before;
            lb_Performance_Referee02.ForeColor = Show_Before;
            lb_Performance_Referee03.ForeColor = Show_Before;
            lb_Performance_Referee04.ForeColor = Show_Before;
            lb_Performance_Referee05.ForeColor = Show_Before;
            lb_Performance_Referee06.ForeColor = Show_Before;
            lb_Performance_Referee07.ForeColor = Show_Before;
            //----------
            lb_Pose_TotalScore.ForeColor = Show_Before;
            //----------
            lb_Correctness_Score.ForeColor = Show_Before;
            lb_Performance_Score.ForeColor = Show_Before;
            //ColorIni///////////////////////////////////////////////////////////////////////

            //
            Count_Referee = comt.Form1.Count_Referee;//有幾位評審
            //

            if (Count_Referee <= 5)
            {
                //TempletIni///////////////////////////////////////////////////////////////////////////
                #region 排版
                //正確性背////////////////////////////////////////////////
                this.panel11.Location = new System.Drawing.Point(66, 9);
                this.panel11.Size = new System.Drawing.Size(164, 146);

                this.panel12.Location = new System.Drawing.Point(66, 156);
                this.panel12.Size = new System.Drawing.Size(164, 146);

                this.panel13.Location = new System.Drawing.Point(66, 302);
                this.panel13.Size = new System.Drawing.Size(164, 146);

                this.panel14.Location = new System.Drawing.Point(66, 448);
                this.panel14.Size = new System.Drawing.Size(164, 146);

                this.panel15.Location = new System.Drawing.Point(66, 594);
                this.panel15.Size = new System.Drawing.Size(164, 146);
                //-----------------------------------------
                this.panel16.Location = new System.Drawing.Point(1127, 315);
                this.panel16.Size = new System.Drawing.Size(164, 104);

                this.panel17.Location = new System.Drawing.Point(1127, 419);
                this.panel17.Size = new System.Drawing.Size(164, 104);
                //正確性背////////////////////////////////////////////////

                //正確性字////////////////////////////////////////////////
                this.lb_Correctness_Referee01.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee01.Location = new System.Drawing.Point(-18, 11);
                this.lb_Correctness_Referee01.Size = new System.Drawing.Size(190, 120);

                this.lb_Correctness_Referee02.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee02.Location = new System.Drawing.Point(-18, 11);
                this.lb_Correctness_Referee02.Size = new System.Drawing.Size(190, 120);

                this.lb_Correctness_Referee03.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee03.Location = new System.Drawing.Point(-18, 9);
                this.lb_Correctness_Referee03.Size = new System.Drawing.Size(190, 120);

                this.lb_Correctness_Referee04.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee04.Location = new System.Drawing.Point(-18, 10);
                this.lb_Correctness_Referee04.Size = new System.Drawing.Size(190, 120);

                this.lb_Correctness_Referee05.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee05.Location = new System.Drawing.Point(-18, 11);
                this.lb_Correctness_Referee05.Size = new System.Drawing.Size(190, 120);

                this.lb_Correctness_Referee06.Font = new System.Drawing.Font("新細明體", 80.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee06.Location = new System.Drawing.Point(-9, -4);
                this.lb_Correctness_Referee06.Size = new System.Drawing.Size(170, 107);

                this.lb_Correctness_Referee07.Font = new System.Drawing.Font("新細明體", 80.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Correctness_Referee07.Location = new System.Drawing.Point(-9, -4);
                this.lb_Correctness_Referee07.Size = new System.Drawing.Size(170, 107);
                //正確性字////////////////////////////////////////////////

                //表現性底////////////////////////////////////////////////
                this.panel18.Location = new System.Drawing.Point(1291, 315);
                this.panel18.Size = new System.Drawing.Size(164, 104);

                this.panel19.Location = new System.Drawing.Point(230, 594);
                this.panel19.Size = new System.Drawing.Size(164, 146);

                this.panel20.Location = new System.Drawing.Point(230, 302);
                this.panel20.Size = new System.Drawing.Size(164, 146);

                this.panel21.Location = new System.Drawing.Point(1291, 419);
                this.panel21.Size = new System.Drawing.Size(164, 104);

                this.panel22.Location = new System.Drawing.Point(230, 448);
                this.panel22.Size = new System.Drawing.Size(164, 146);

                this.panel23.Location = new System.Drawing.Point(230, 156);
                this.panel23.Size = new System.Drawing.Size(164, 146);

                this.panel24.Location = new System.Drawing.Point(230, 9);
                this.panel24.Size = new System.Drawing.Size(164, 146);
                //表現性底////////////////////////////////////////////////

                //表現性字////////////////////////////////////////////////
                this.lb_Performance_Referee01.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee01.Location = new System.Drawing.Point(-21, 11);
                this.lb_Performance_Referee01.Size = new System.Drawing.Size(190, 120);

                this.lb_Performance_Referee02.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee02.Location = new System.Drawing.Point(-21, 11);
                this.lb_Performance_Referee02.Size = new System.Drawing.Size(190, 120);

                this.lb_Performance_Referee03.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee03.Location = new System.Drawing.Point(-21, 11);
                this.lb_Performance_Referee03.Size = new System.Drawing.Size(190, 120);

                this.lb_Performance_Referee04.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee04.Location = new System.Drawing.Point(-21, 10);
                this.lb_Performance_Referee04.Size = new System.Drawing.Size(190, 120);

                this.lb_Performance_Referee05.Font = new System.Drawing.Font("新細明體", 90F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee05.Location = new System.Drawing.Point(-21, 11);
                this.lb_Performance_Referee05.Size = new System.Drawing.Size(190, 120);

                this.lb_Performance_Referee06.Font = new System.Drawing.Font("新細明體", 80.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee06.Location = new System.Drawing.Point(-8, -4);
                this.lb_Performance_Referee06.Size = new System.Drawing.Size(170, 107);

                this.lb_Performance_Referee07.Font = new System.Drawing.Font("新細明體", 80.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Performance_Referee07.Location = new System.Drawing.Point(-8, -4);
                this.lb_Performance_Referee07.Size = new System.Drawing.Size(170, 107);
                //表現性字////////////////////////////////////////////////

                //裁判底//////////////////////////////////////////////////
                this.panel4.Location = new System.Drawing.Point(2, 9);
                this.panel4.Size = new System.Drawing.Size(70, 146);

                this.panel5.Location = new System.Drawing.Point(2, 155);
                this.panel5.Size = new System.Drawing.Size(70, 146);

                this.panel6.Location = new System.Drawing.Point(2, 301);
                this.panel6.Size = new System.Drawing.Size(70, 146);

                this.panel7.Location = new System.Drawing.Point(2, 447);
                this.panel7.Size = new System.Drawing.Size(70, 146);

                this.panel8.Location = new System.Drawing.Point(2, 594);
                this.panel8.Size = new System.Drawing.Size(70, 146);

                this.panel9.Location = new System.Drawing.Point(1059, 315);
                this.panel9.Size = new System.Drawing.Size(70, 104);

                this.panel10.Location = new System.Drawing.Point(1059, 419);
                this.panel10.Size = new System.Drawing.Size(70, 104);
                //裁判底//////////////////////////////////////////////////

                //裁判字//////////////////////////////////////////////////
                this.lb_Referee01_Title.Font = new System.Drawing.Font("新細明體", 60F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee01_Title.Location = new System.Drawing.Point(-2, 30);
                this.lb_Referee01_Title.Size = new System.Drawing.Size(72, 80);

                this.lb_Referee02_Title.Font = new System.Drawing.Font("新細明體", 60F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee02_Title.Location = new System.Drawing.Point(-2, 30);
                this.lb_Referee02_Title.Size = new System.Drawing.Size(72, 80);

                this.lb_Referee03_Title.Font = new System.Drawing.Font("新細明體", 60F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee03_Title.Location = new System.Drawing.Point(-2, 30);
                this.lb_Referee03_Title.Size = new System.Drawing.Size(72, 80);

                this.lb_Referee04_Title.Font = new System.Drawing.Font("新細明體", 60F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee04_Title.Location = new System.Drawing.Point(-2, 33);
                this.lb_Referee04_Title.Size = new System.Drawing.Size(72, 80);

                this.lb_Referee05_Title.Font = new System.Drawing.Font("新細明體", 60F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee05_Title.Location = new System.Drawing.Point(-2, 33);
                this.lb_Referee05_Title.Size = new System.Drawing.Size(72, 80);

                this.lb_Referee06_Title.Font = new System.Drawing.Font("新細明體", 45F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee06_Title.Location = new System.Drawing.Point(3, 19);
                this.lb_Referee06_Title.Size = new System.Drawing.Size(53, 60);

                this.lb_Referee07_Title.Font = new System.Drawing.Font("新細明體", 45F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lb_Referee07_Title.Location = new System.Drawing.Point(3, 20);
                this.lb_Referee07_Title.Size = new System.Drawing.Size(53, 60);
                //裁判字//////////////////////////////////////////////////
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
                lb_Performance_Referee01.Text = "";
            }
            if (2 > Count_Referee)
            {
                lb_Performance_Referee02.Text = "";
            }
            if (3 > Count_Referee)
            {
                lb_Performance_Referee03.Text = "";
            }
            if (4 > Count_Referee)
            {
                lb_Performance_Referee04.Text = "";
            }
            if (5 > Count_Referee)
            {
                lb_Performance_Referee05.Text = "";
            }
            if (6 > Count_Referee)
            {
                lb_Performance_Referee06.Text = "";
            }
            if (7 > Count_Referee)
            {
                lb_Performance_Referee07.Text = "";
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
                lb_NowPose.Text = "第一品勢";//lb_NowPose
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
                lb_NowPose.Text = "第二品勢";//lb_NowPose
                //-----------------------------------------------------------------------------------
                Select_Pose = comt.Form1.Select_Pose;//這次比賽是全部都單品勢還是全部都雙品勢
            }
        }

        private void timer_Rating_Tick(object sender, EventArgs e)
        {
            Flag_Correctness = comt.Form1.Flag_Correctness;
            Flag_Performance = comt.Form1.Flag_Performance;
            Score_Correctness = comt.Form1.Score_Correctness;
            Score_Performance = comt.Form1.Score_Performance;
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
                    lb_Correctness_Referee01.Text = Score_Correctness[1];
                    if (("1" == Lowest_Score_Correctness || "1" == Highest_Score_Correctness) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Correctness_Referee01.ForeColor = Color.Red;
                    }
                }
                if (2 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee02.ForeColor = Show_After;
                    lb_Correctness_Referee02.Text = Score_Correctness[2];
                    if (("2" == Lowest_Score_Correctness || "2" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee02.ForeColor = Color.Red;
                    }
                }
                if (3 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee03.ForeColor = Show_After;
                    lb_Correctness_Referee03.Text = Score_Correctness[3];
                    if (("3" == Lowest_Score_Correctness || "3" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee03.ForeColor = Color.Red;
                    }
                }
                if (4 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee04.ForeColor = Show_After;
                    lb_Correctness_Referee04.Text = Score_Correctness[4];
                    if (("4" == Lowest_Score_Correctness || "4" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee04.ForeColor = Color.Red;
                    }
                }
                if (5 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee05.ForeColor = Show_After;
                    lb_Correctness_Referee05.Text = Score_Correctness[5];
                    if (("5" == Lowest_Score_Correctness || "5" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee05.ForeColor = Color.Red;
                    }
                }
                if (6 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee06.ForeColor = Show_After;
                    lb_Correctness_Referee06.Text = Score_Correctness[6];
                    if (("6" == Lowest_Score_Correctness || "6" == Highest_Score_Correctness) && Count_Referee > 3)
                    {
                        lb_Correctness_Referee06.ForeColor = Color.Red;
                    }
                }
                if (7 <= Count_Referee && count_Correctness_AlreadyRating == Count_Referee)
                {
                    lb_Correctness_Referee07.ForeColor = Show_After;
                    lb_Correctness_Referee07.Text = Score_Correctness[7];
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
                    lb_Correctness_Referee01.ForeColor = Show_After;
                    lb_Correctness_Referee01.Text = "OK";
                }
                if (Flag_Correctness[2] == "1" && 2 <= Count_Referee)
                {
                    lb_Correctness_Referee02.ForeColor = Show_After;
                    lb_Correctness_Referee02.Text = "OK";
                }
                if (Flag_Correctness[3] == "1" && 3 <= Count_Referee)
                {
                    lb_Correctness_Referee03.ForeColor = Show_After;
                    lb_Correctness_Referee03.Text = "OK";
                }
                if (Flag_Correctness[4] == "1" && 4 <= Count_Referee)
                {
                    lb_Correctness_Referee04.ForeColor = Show_After;
                    lb_Correctness_Referee04.Text = "OK";
                }
                if (Flag_Correctness[5] == "1" && 5 <= Count_Referee)
                {
                    lb_Correctness_Referee05.ForeColor = Show_After;
                    lb_Correctness_Referee05.Text = "OK";
                }
                if (Flag_Correctness[6] == "1" && 6 <= Count_Referee)
                {
                    lb_Correctness_Referee06.ForeColor = Show_After;
                    lb_Correctness_Referee06.Text = "OK";
                }
                if (Flag_Correctness[7] == "1" && 7 <= Count_Referee)
                {
                    lb_Correctness_Referee07.ForeColor = Show_After;
                    lb_Correctness_Referee07.Text = "OK";
                }
            }

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

                Double_Score_Average_Performance = 0;//初始化
                //算表現性分數平均/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (Count_Referee == 1)//如果只有一個人,最低分就是平均分數，ie,sort完以後的位置[1]
                {
                    Double_Score_Average_Performance = ScoreRank_Performance_Variable[1].Score;
                    String_Score_Average_Performance = Double_Score_Average_Performance.ToString("0.00");//String_Score_Average_Performance = Lowest_Score_Performance;
                }
                else if (Count_Referee == 2)//如果有兩個人，就( 位置[1] +位置[2] )在除以 2
                {
                    Double_Score_Average_Performance = ((double)((ScoreRank_Performance_Variable[1].Score + ScoreRank_Performance_Variable[2].Score) / 2));
                    String_Score_Average_Performance = Double_Score_Average_Performance.ToString("0.00");//.ToString("0.00")代表四捨五入到小數點後面第2位
                }
                else if (Count_Referee == 3)
                {
                    Double_Score_Average_Performance = ((double)((ScoreRank_Performance_Variable[1].Score + ScoreRank_Performance_Variable[2].Score + ScoreRank_Performance_Variable[3].Score) / 3));
                    String_Score_Average_Performance = Double_Score_Average_Performance.ToString("0.00");
                }
                else if (Count_Referee > 3)//如果有三個人以上，就從位置[2] 加到 位置[最高位-1] 因為要 去掉"最高" 和 "最低"(ie,位置0)
                {
                    for (int i = 2; i <= (Count_Referee - 1); i++)
                    {
                        Double_Score_Average_Performance = Double_Score_Average_Performance + ScoreRank_Performance_Variable[i].Score;
                    }
                    Double_Score_Average_Performance = Double_Score_Average_Performance / (Count_Referee - 2);
                    String_Score_Average_Performance = Double_Score_Average_Performance.ToString("0.00");
                }
                lb_Performance_Score.ForeColor = Show_After;
                lb_Performance_Score.Text = String_Score_Average_Performance;
                //算表現性分數平均/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                Double_TotalScore_Correctness_Plus_Performance = Double_Score_Average_Correctness + Double_Score_Average_Performance;
                String_TotalScore_Correctness_Plus_Performance = Double_TotalScore_Correctness_Plus_Performance.ToString("0.00");

                //全部分數Data只在這裡進行寫入///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (Now_Pose == 1)
                {
                    comt.Form1.ImportData[6, comt.Form1.Now_Player] = String_Score_Average_Correctness;//正當性平均分數                    
                    comt.Form1.ImportData[7, comt.Form1.Now_Player] = String_Score_Average_Performance;//表現性平均分數
                    comt.Form1.ImportData[8, comt.Form1.Now_Player] = String_TotalScore_Correctness_Plus_Performance;//正當性+表現性總分
                    lb_Pose_TotalScore.ForeColor = Show_After;
                    lb_Pose_TotalScore.Text = String_TotalScore_Correctness_Plus_Performance;
                    timer_ForCloseWinForm.Enabled = true;
                    //IsRating = true;
                    /*
                    Thread.Sleep(5000);
                    this.Close();
                    */
                }
                if (Now_Pose == 2)
                {
                    comt.Form1.ImportData[9, comt.Form1.Now_Player] = String_Score_Average_Correctness;//正當性平均分數
                    comt.Form1.ImportData[10, comt.Form1.Now_Player] = String_Score_Average_Performance;//表現性平均分數
                    comt.Form1.ImportData[11, comt.Form1.Now_Player] = String_TotalScore_Correctness_Plus_Performance;//正當性+表現性總分
                    lb_Pose_TotalScore.ForeColor = Show_After;
                    lb_Pose_TotalScore.Text = String_TotalScore_Correctness_Plus_Performance;
                    timer_ForCloseWinForm.Enabled = true;
                    //IsRating = true;
                    /*
                    Thread.Sleep(5000);
                    this.Close();
                    */
                }
                //全部分數Data只在這裡進行寫入///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (1 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)//看有幾位評審，就翻幾個牌，因為翻1號評審，所以拿1去跟評審個數比，若不大於就翻牌
                {
                    lb_Performance_Referee01.ForeColor = Show_After;
                    lb_Performance_Referee01.Text = Score_Performance[1];
                    if (("1" == Lowest_Score_Performance || "1" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee01.ForeColor = Color.Red;
                    }
                }
                if (2 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee02.ForeColor = Show_After;
                    lb_Performance_Referee02.Text = Score_Performance[2];
                    if (("2" == Lowest_Score_Performance || "2" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee02.ForeColor = Color.Red;
                    }
                }
                if (3 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee03.ForeColor = Show_After;
                    lb_Performance_Referee03.Text = Score_Performance[3];
                    if (("3" == Lowest_Score_Performance || "3" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee03.ForeColor = Color.Red;
                    }
                }
                if (4 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee04.ForeColor = Show_After;
                    lb_Performance_Referee04.Text = Score_Performance[4];
                    if (("4" == Lowest_Score_Performance || "4" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee04.ForeColor = Color.Red;
                    }
                }
                if (5 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee05.ForeColor = Show_After;
                    lb_Performance_Referee05.Text = Score_Performance[5];
                    if (("5" == Lowest_Score_Performance || "5" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee05.ForeColor = Color.Red;
                    }
                }
                if (6 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee06.ForeColor = Show_After;
                    lb_Performance_Referee06.Text = Score_Performance[6];
                    if (("6" == Lowest_Score_Performance || "6" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee06.ForeColor = Color.Red;
                    }
                }
                if (7 <= Count_Referee && count_Performance_AlreadyRating == Count_Referee)
                {
                    lb_Performance_Referee07.ForeColor = Show_After;
                    lb_Performance_Referee07.Text = Score_Performance[7];
                    if (("7" == Lowest_Score_Performance || "7" == Highest_Score_Performance) && Count_Referee > 3)//如果最高分或最低分是"1"號評審給的，就將他反白，不予計分 且 要評審數3個以上才使用 最高最低評分機制
                    {
                        lb_Performance_Referee07.ForeColor = Color.Red;
                    }
                }
            }
            else
            {
                //if (count_Correctness_AlreadyRating == Count_Referee)//正確顯示完再顯示表現
                //{
                    if (Flag_Performance[1] == "1" && 1 <= Count_Referee)
                    {
                        lb_Performance_Referee01.ForeColor = Show_After;
                        lb_Performance_Referee01.Text = "OK";
                    }
                    if (Flag_Performance[2] == "1" && 2 <= Count_Referee)
                    {
                        lb_Performance_Referee02.ForeColor = Show_After;
                        lb_Performance_Referee02.Text = "OK";
                    }
                    if (Flag_Performance[3] == "1" && 3 <= Count_Referee)
                    {
                        lb_Performance_Referee03.ForeColor = Show_After;
                        lb_Performance_Referee03.Text = "OK";
                    }
                    if (Flag_Performance[4] == "1" && 4 <= Count_Referee)
                    {
                        lb_Performance_Referee04.ForeColor = Show_After;
                        lb_Performance_Referee04.Text = "OK";
                    }
                    if (Flag_Performance[5] == "1" && 5 <= Count_Referee)
                    {
                        lb_Performance_Referee05.ForeColor = Show_After;
                        lb_Performance_Referee05.Text = "OK";
                    }
                    if (Flag_Performance[6] == "1" && 6 <= Count_Referee)
                    {
                        lb_Performance_Referee06.ForeColor = Show_After;
                        lb_Performance_Referee06.Text = "OK";
                    }
                    if (Flag_Performance[7] == "1" && 7 <= Count_Referee)
                    {
                        lb_Performance_Referee07.ForeColor = Show_After;
                        lb_Performance_Referee07.Text = "OK";
                    }
                //}
                
            }
        }
        public static Result Result;
        public static Rank Rank;
        private void Rating_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (Select_Pose == "1" && Now_Pose == 1)//只有在 只比一個品勢 且一個品勢 已經比完 才會開啟結果視窗
            {
                /*
                //開啟結果前將結果寫入總分
                comt.Form1.ImportData[12, comt.Form1.Now_Player] = comt.Form1.ImportData[8, comt.Form1.Now_Player];//因為單一品勢，所以 總平均 = 單一品勢分數
                Result = new Result();
                Result.Show();
                */
                comt.Form1.ImportData[12, comt.Form1.Now_Player] = comt.Form1.ImportData[8, comt.Form1.Now_Player];//因為單一品勢，所以 總平均 = 單一品勢分數
                comt.Form1.WriteExcelAll();
                comt.Form1.IsOpen_Rank_ByRating = true;
                Rank = new Rank();
                Rank.Show();
                //Thread.Sleep(50);                
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
                comt.Form1.ImportData[12, comt.Form1.Now_Player] = ((System.Convert.ToDouble(comt.Form1.ImportData[8, comt.Form1.Now_Player]) + System.Convert.ToDouble(comt.Form1.ImportData[11, comt.Form1.Now_Player])) / 2).ToString("0.00");
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

