using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//--------------------------------------
using RankShow;

namespace comt
{

    public partial class Rank : Form
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

        //變數//////////////////////////////////////////////////////////////////////////////
        int Sum_ASCII_Number = 0;//編號的ascii相加
        double Final_Weighted_Score = 0;//最終加權分數
        //變數//////////////////////////////////////////////////////////////////////////////
        /*
        //排名////////////////////////////////////////////////////////////////////////////
        string[] RealRank_Place;
        string[] RealRank_ID;
        string[] RealRank_Score;
         * 
         * 
         * 
        //排名////////////////////////////////////////////////////////////////////////////
        */
        /*
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;

                return cp;
            }
        }
        */

        //排名//////////////////////////////////////////////
        int RankPlace = 0;
        int SamePlace = 0;
        //排名//////////////////////////////////////////////

        //Location///////////////////////////////////////////////////
        int Location_Bar_X = 44;
        int Location_Bar_Y_List1= 139;
        int Location_Bar_Y_List2 = 243;
        int Location_Bar_Y_List3 = 350;
        int Location_Bar_Y_List4 = 455;
        int Location_Bar_Y_List5 = 562;
        int Location_Bar_Y_List6 = 662;
        //Location///////////////////////////////////////////////////
        class ScoreRank_Total
        {
            public string Rank;//名次
            public string Number;//編號
            public string Name;//姓名
            public double Score;//分數
            public int ArrayPosition;//陣列位置
            public string RealPlace;//陣列位置

            public ScoreRank_Total(string Rank, string Number, string Name, double Score, int ArrayPosition, string RealPlace)
            {
                this.Rank = Rank;
                this.Number = Number;
                this.Name = Name;
                this.Score = Score;
                this.ArrayPosition = ArrayPosition;
                this.RealPlace = RealPlace;
            }
        }
        ScoreRank_Total[] ScoreRank_Total_Variable;
        private int SortMethod_Total(ScoreRank_Total ScoreRank_Total1, ScoreRank_Total ScoreRank_Total2)
        {
            return ScoreRank_Total1.Score.CompareTo(ScoreRank_Total2.Score);
        }
        int Real_AlreadyRatedPlayerCount = 0;//Update_20140518_01
        public Rank()
        {
            /*
            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true); // 禁止擦除背景.
            SetStyle(ControlStyles.OptimizedDoubleBuffer, true); // 双缓冲
            */
            //Excel////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //0:組別
            //1:單位
            //2:姓名
            //3:初賽或複賽
            //4:第一品勢名稱
            //5:第二品勢名稱 
            //6:第一品勢正確性分數
            //7:第一品勢表現性分數
            //8:第一品勢總分
            //9:第二品勢正確性分數
            //10:第二品勢表現性分數
            //11:第二品勢總分
            //12:總平均
            //13:排名
            //14:名次
            //15:編號
            //16.棄權
            //Excel////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //SetStyle(ControlStyles.UserPaint, true);
            //SetStyle(ControlStyles.AllPaintingInWmPaint, true); // 禁止擦除背景.
            //SetStyle(ControlStyles.OptimizedDoubleBuffer, true); // 双缓冲
            //移動視窗到延伸桌面
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);


            //顯示: 1.名次 2.編號 3.姓名 4.分數 (額外:5陣列位置)
            InitializeComponent();

            comt.Form1.IsWritng_Rank = true;//正在寫入
            //ini///////////////////////////////////////////////////////////////////
            ImportData = comt.ImportSetting.ImportData;//把exl資料表格載入
            Select_Pose = comt.ImportSetting.Select_Pose;//在ImportSetting時選的要比幾個品勢
            Raw_CompetitorNumber = comt.ImportSetting.Raw_CompetitorNumber;//共有多少位參賽者
            Column_Pose1 = comt.ImportSetting.Column_Pose1;//若選單品勢會有幾行
            Column_Pose2 = comt.ImportSetting.Column_Pose2;//若選雙品勢會有幾行
            Count_Referee = comt.ImportSetting.Count_Referee;//有幾位評審
            ImportData_Score = comt.ImportSetting.ImportData_Score;//把exl分數表格載入    
            //ini///////////////////////////////////////////////////////////////////
            Now_Pose = comt.Form1.Now_Pose;
            Now_Player = comt.Form1.Now_Player;

            //Update_20140518_02_原本沒有_多加入，若回來重頻棄權的人時，才可以知道之前頻道第幾個人，再一起做排序/////////////////////
            Real_AlreadyRatedPlayerCount = 0;
            for (int i = 1; i <= Raw_CompetitorNumber; i++)
            {
                if (ImportData[21, i].ToString() != "")
                {
                    Real_AlreadyRatedPlayerCount++;
                }
            }
            //Update_20140518/////////////////////
            //ini/////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region ini
            lb_Rank01.Text = "";
            lb_Rank02.Text = "";
            lb_Rank03.Text = "";
            lb_Rank04.Text = "";
            lb_Rank05.Text = "";
            lb_Rank06.Text = "";
            lb_Rank07.Text = "";
            lb_Rank08.Text = "";
            //---------------------------------------------------------------------
            lb_ID_01.Text = "";
            lb_ID_02.Text = "";
            lb_ID_03.Text = "";
            lb_ID_04.Text = "";
            lb_ID_05.Text = "";
            lb_ID_06.Text = "";
            lb_ID_07.Text = "";
            lb_ID_08.Text = "";
            //---------------------------------------------------------------------
            lb_Name_01.Text = "";
            lb_Name_02.Text = "";
            lb_Name_03.Text = "";
            lb_Name_04.Text = "";
            lb_Name_05.Text = "";
            lb_Name_06.Text = "";
            lb_Name_07.Text = "";
            lb_Name_08.Text = "";
            //---------------------------------------------------------------------
            lb_Score_01.Text = "";
            lb_Score_02.Text = "";
            lb_Score_03.Text = "";
            lb_Score_04.Text = "";
            lb_Score_05.Text = "";
            lb_Score_06.Text = "";
            lb_Score_07.Text = "";
            lb_Score_08.Text = "";
            #endregion

            #region 透明
            //this.pictureBox1.Controls.Add(tblp_Rank1);
            /*
            tblp_Rank1.Parent = pictureBox1;
            tblp_Rank1.Visible = true;
            this.tblp_Rank1.Location = new Point(44, 139);
            this.lb_Rank01.Location = new Point(44, 139);

            tblp_Rank1.BringToFront();
            tblp_Rank2.Parent = pictureBox2;
            tblp_Rank3.Parent = pictureBox3;
            tblp_Rank4.Parent = pictureBox4;
            tblp_Rank5.Parent = pictureBox5;
            tblp_Rank6.Parent = pictureBox6;
            
            tblp_Number1.Parent = pictureBox1;
            tblp_Number2.Parent = pictureBox2;
            tblp_Number3.Parent = pictureBox3;
            tblp_Number4.Parent = pictureBox4;
            tblp_Number5.Parent = pictureBox5;
            tblp_Number6.Parent = pictureBox6;

            tblp_Score1.Parent = pictureBox1;
            tblp_Score2.Parent = pictureBox2;
            tblp_Score3.Parent = pictureBox3;
            tblp_Score4.Parent = pictureBox4;
            tblp_Score5.Parent = pictureBox5;
            tblp_Score6.Parent = pictureBox6;

            tblp_Name1.Parent = pictureBox1;
            tblp_Name2.Parent = pictureBox2;
            tblp_Name3.Parent = pictureBox3;
            tblp_Name4.Parent = pictureBox4;
            tblp_Name5.Parent = pictureBox5;
            tblp_Name6.Parent = pictureBox6;
            */
            #endregion
            //ini/////////////////////////////////////////////////////////////////////////////////////////////////////////
            ScoreRank_Total_Variable = new ScoreRank_Total[Real_AlreadyRatedPlayerCount + 1];//Update_20140518_03//ScoreRank_Total_Variable = new ScoreRank_Total[Now_Player + 1];
            for (int i = 0; i <= Real_AlreadyRatedPlayerCount; i++)//Update_20140518_04_for (int i = 0; i <= Now_Player; i++)
            {
                if (i == 0)
                {
                    ScoreRank_Total_Variable[0] = new ScoreRank_Total("Rank", "Number", "Name", -100000000000.0, 0,"");//因為後面有加權，要-10*10^12才夠小
                }
                else 
                {
                    //算編號的ADCII總和///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    Sum_ASCII_Number = 0;
                    
                    for (int j = 0; j < comt.Form1.ImportData[15, i].Length; j++)
                    { 
                        string String_Substring = comt.Form1.ImportData[15, i].Substring(j,1);
                        string String_Substring_To_HexASCII = Convert.ToString(ASC(String_Substring), 16);
                        int Int_HexASCII_To_Decimal = Int32.Parse(String.Format("{0:X}", String_Substring_To_HexASCII), System.Globalization.NumberStyles.HexNumber);
                        Sum_ASCII_Number = Sum_ASCII_Number + Int_HexASCII_To_Decimal;
                        //comt.Form1.ImportData[15, i].Substring(j,1);//16toDecimal(ASCII(substring))
                        //Sum_ASCII_Number = Sum_ASCII_Number + ;
                    }
                    //算編號的ADCII總和///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    if (Select_Pose == "1") 
                    {
                        string temp_asda = comt.Form1.ImportData[12, i].ToString();
                        string temp_asfa = comt.Form1.ImportData[7, i].ToString();
                        //原本Final_Weighted_Score = System.Convert.ToDouble(comt.Form1.ImportData[12, i]) * 1000000000 + System.Convert.ToDouble(comt.Form1.ImportData[7, i]) * 1000000;// +Sum_ASCII_Number;
                        //總分*1000000000 + 第一表現性平均*1000000 + (選手編號ASCII相加)
                        Final_Weighted_Score = System.Convert.ToDouble(comt.Form1.ImportData[12, i]) * 1000000000 + System.Convert.ToDouble(comt.Form1.ImportData[7, i]) * 1000000 + System.Convert.ToDouble(comt.Form1.ImportData[21, i]);// +Sum_ASCII_Number;
                       //總分*1000000000 + 第一表現性平均*1000000 + (選手編號ASCII相加) + 原始總分
                    }
                    if (Select_Pose == "2")
                    {
                        //原本Final_Weighted_Score = System.Convert.ToDouble(comt.Form1.ImportData[12, i]) * 1000000000 + ((System.Convert.ToDouble(comt.Form1.ImportData[7, i]) + System.Convert.ToDouble(comt.Form1.ImportData[10, comt.Form1.Now_Player])) / 2) * 1000000;// +Sum_ASCII_Number;
                        //總分*1000000000 + ((第一表現性平均+第二表現性平均)/2)*1000000 + (選手編號ASCII相加)
                        Final_Weighted_Score = System.Convert.ToDouble(comt.Form1.ImportData[12, i]) * 1000000000 + ((System.Convert.ToDouble(comt.Form1.ImportData[7, i]) + System.Convert.ToDouble(comt.Form1.ImportData[10, i])) / 2) * 1000000 + System.Convert.ToDouble(comt.Form1.ImportData[21, i]);// +Sum_ASCII_Number;
                        //總分*1000000000 + ((第一表現性平均+第二表現性平均)/2)*1000000 + (選手編號ASCII相加) + 原始總分
                    }
                     
                    ScoreRank_Total_Variable[i] = new ScoreRank_Total
                    (
                        comt.Form1.ImportData[12, i],//總平均
                        comt.Form1.ImportData[15, i],//Number
                        comt.Form1.ImportData[2, i], //Name
                        Final_Weighted_Score,//System.Convert.ToDouble(comt.Form1.ImportData[12, i]),//Score
                        i,//ArrayPosition
                        ""
                    );                    
                }
            }

                

                /*
                //寫入名次//
                for (i = 1; i <= Now_Player; i++)
                {
                    comt.Form1.ImportData[13, i];
                }
                */
                ///////////加入中文字排名

            Array.Sort(ScoreRank_Total_Variable, SortMethod_Total);
            int Counter_Quit = 0;
            /*
            RealRank_Place = new string[Now_Player + 1];
            RealRank_ID = new string[Now_Player + 1];
            RealRank_Score = new string[Now_Player + 1];
            */

            //將名次填入陣列中/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            for (int k = Real_AlreadyRatedPlayerCount; k >= 1; k--)//Update_20140518_05_for (int k = Now_Player; k >= 1 ; k--)
            {
                int position = ScoreRank_Total_Variable[k].ArrayPosition;//position是原本陣列中的位置
                //GetChineseNumber(int number)
                if (comt.Form1.ImportData[16, position] == "1")//若找到棄權的人
                {
                    Counter_Quit++;//棄權的人數再加1
                    comt.Form1.ImportData[13, position] = "";//名次為空，不計名次
                    comt.Form1.ImportData[14, position] = "";//國字名次
                }
                else 
                {
                    if (ScoreRank_Total_Variable.Length - 1 > k)
                    {
                        if (ScoreRank_Total_Variable[k].Score == ScoreRank_Total_Variable[k + 1].Score)//和上一位同分
                        {
                            SamePlace++;
                            comt.Form1.ImportData[13, position] = (Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit - SamePlace).ToString();//Update_20140518_06_comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            ScoreRank_Total_Variable[k].RealPlace = (Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit - SamePlace).ToString();//Update_20140518_07_ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            //---------------------------
                            comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit - SamePlace) + "名";//Update_20140518_08_國字名次//comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Now_Player - k + 1 - Counter_Quit - SamePlace) + "名";//國字名次
                            /*
                            int past = ScoreRank_Total_Variable[k+1].ArrayPosition;
                            comt.Form1.ImportData[13, past] = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            ScoreRank_Total_Variable[k + 1].RealPlace = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            */
                        }
                        if (ScoreRank_Total_Variable[k].Score < ScoreRank_Total_Variable[k + 1].Score)//分數比上一位低
                        {
                            SamePlace = 0;
                            //1223
                            /*
                            comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            */
                            //1224
                            comt.Form1.ImportData[13, position] = (Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit).ToString();//Update_20140518_09_comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit).ToString();
                            ScoreRank_Total_Variable[k].RealPlace = (Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit).ToString();//Update_20140518_10_ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit).ToString();
                            //---------------------------
                            comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit) + "名";//Update_20140518_11_國字名次//comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Now_Player - k + 1 - Counter_Quit) + "名";//國字名次
                        }
                        //用原本陣列中的位置 = 給他名次 ， 從最高的名次給 到 最低
                    }
                    else
                    {
                        comt.Form1.ImportData[13, position] = (Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit - SamePlace).ToString();//Update_20140518_12_第一名//comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();//第一名                       
                        ScoreRank_Total_Variable[k].RealPlace = (Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit - SamePlace).ToString();//Update_20140518_13_ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                        //---------------------------
                        comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Real_AlreadyRatedPlayerCount - k + 1 - Counter_Quit - SamePlace) + "名";//Update_20140518_14_國字名次//comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Now_Player - k + 1 - Counter_Quit - SamePlace)+"名";//國字名次
                    }                    
                }

            }
            //將名次填入陣列中/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            timer_bling.Enabled = true;

            #region mark
            /*
            string k1 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8].Number;
            string k2 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8].Name;
            string k2_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8].Rank;
            string k3 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 1].Number;
            string k4 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 1].Name;
            string k4_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 1].Rank;
            string k11 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8-2].Number;
            string k21 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8-2].Name;
            string k21_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8-2].Rank;
            string k31 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 3].Number;
            string k41 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 3].Name;
            string k41_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 3].Rank;
            */
            /*
            for (int i = 0; i <= Now_Player; i++)
            {
                textBox1.AppendText(ScoreRank_Total_Variable[i].Name + "  " +ScoreRank_Total_Variable[i].Score + "\n");
            }
            */

            /*
            #region Display

            ////找出現在選手的名次///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            int Now_Player_Place = 1;//找出現在選手的名次
            for (int m = Now_Player; m >= 1; m--)
            {
                if (comt.Form1.ImportData[2, Now_Player] == ScoreRank_Total_Variable[m].Name && comt.Form1.ImportData[15, Now_Player] == ScoreRank_Total_Variable[m].Number)
                {
                    Now_Player_Place = Now_Player - m + 1;
                }
            }
            ////找出現在選手的名次///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //顯示/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            int Display_Place;//顯示名次的位置
            int Display_Page;//顯示名次的頁碼
            Display_Place = Now_Player_Place % 6;
            if (Display_Place == 0)
            {
                Display_Page = (Now_Player_Place / 6) - 1;
            }
            else 
            {
                Display_Page = Now_Player_Place / 6;
            }
            


                if (Display_Place == 1)
                {
                    if (Now_Player - Display_Page * 6 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                    {
                        lb_Rank01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].RealPlace; //(1 + Display_Page * 8).ToString();
                        //lb_Rank01.ForeColor = Color.Black;
                        lb_ID_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Number;
                        //lb_ID_01.ForeColor = Color.Black;
                        lb_Name_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Name;
                        //lb_Name_01.ForeColor = Color.Black;
                        lb_Score_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Rank;
                        //lb_Score_01.ForeColor = Color.Black;
                        panel1.BackColor = Color.Red;
                        panel9.BackColor = Color.Yellow;
                        pnl_list01.BackgroundImage = comt.Properties.Resources.listred;
                        
                    }
                    else
                    {
                        lb_Rank01.Text = "";
                        lb_ID_01.Text = "";
                        lb_Name_01.Text = "";
                        lb_Score_01.Text = "";
                    }
                }
                else
                {
                    if (Now_Player - Display_Page * 6 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                    {
                        if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Rank == "-1")
                        {
                            lb_Rank01.Text = "";
                            lb_ID_01.Text = "";
                            lb_Name_01.Text = "";
                            lb_Score_01.Text = "";
                        }
                        else 
                        {
                            lb_Rank01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].RealPlace;
                            lb_ID_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Number;
                            lb_Name_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Name;
                            lb_Score_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Rank;
                        }                       
                    }
                    else
                    {
                        lb_Rank01.Text = "";
                        lb_ID_01.Text = "";
                        lb_Name_01.Text = "";
                        lb_Score_01.Text = "";
                    }
                }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 2)
            {
                if (Now_Player - Display_Page * 6 -1>= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].RealPlace;
                    //lb_Rank02.ForeColor = Color.Black;
                    lb_ID_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Number;
                    //lb_ID_02.ForeColor = Color.Black;
                    lb_Name_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Name;
                    //lb_Name_02.ForeColor = Color.Black;
                    lb_Score_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Rank;
                    //lb_Score_02.ForeColor = Color.Black;
                    panel2.BackColor = Color.Red;
                    panel10.BackColor = Color.Yellow;
                    pnl_list02.BackgroundImage = comt.Properties.Resources.listred;
                }
                else
                {
                    lb_Rank02.Text = "";
                    lb_ID_02.Text = "";
                    lb_Name_02.Text = "";
                    lb_Score_02.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 1 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Rank =="-1")
                    {
                        lb_Rank02.Text = "";
                        lb_ID_02.Text = "";
                        lb_Name_02.Text = "";
                        lb_Score_02.Text = "";
                    }
                    else
                    {
                        lb_Rank02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].RealPlace;
                        lb_ID_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Number;
                        lb_Name_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Name;
                        lb_Score_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Rank;
                    }
                    
                }
                else
                {
                    lb_Rank02.Text = "";
                    lb_ID_02.Text = "";
                    lb_Name_02.Text = "";
                    lb_Score_02.Text = "";
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 3)
            {
                if (Now_Player - Display_Page * 6 - 2 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].RealPlace;
                    //lb_Rank03.ForeColor = Color.Black;
                    lb_ID_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Number;
                    //lb_ID_03.ForeColor = Color.Black;
                    lb_Name_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Name;
                    //lb_Name_03.ForeColor = Color.Black;
                    lb_Score_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Rank;
                    //lb_Score_03.ForeColor = Color.Black;
                    panel3.BackColor = Color.Red;
                    panel11.BackColor = Color.Yellow;
                    pnl_list03.BackgroundImage = comt.Properties.Resources.listred;
                }
                else
                {
                    lb_Rank03.Text = "";
                    lb_ID_03.Text = "";
                    lb_Name_03.Text = "";
                    lb_Score_03.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 2 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Rank=="-1")
                    {
                        lb_Rank03.Text = "";
                        lb_ID_03.Text = "";
                        lb_Name_03.Text = "";
                        lb_Score_03.Text = "";
                    }
                    else
                    {
                        lb_Rank03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].RealPlace;
                        lb_ID_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Number;
                        lb_Name_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Name;
                        lb_Score_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Rank;
                    }
                    
                }
                else
                {
                    lb_Rank03.Text = "";
                    lb_ID_03.Text = "";
                    lb_Name_03.Text = "";
                    lb_Score_03.Text = "";
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 4)
            {
                if (Now_Player - Display_Page * 6 - 3 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].RealPlace;
                    //lb_Rank04.ForeColor = Color.Black;
                    lb_ID_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Number;
                    //lb_ID_04.ForeColor = Color.Black;
                    lb_Name_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Name;
                    //lb_Name_04.ForeColor = Color.Black;
                    lb_Score_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Rank;
                    //lb_Score_04.ForeColor = Color.Black;
                    panel4.BackColor = Color.Red;
                    panel12.BackColor = Color.Yellow;
                    pnl_list04.BackgroundImage = comt.Properties.Resources.listred;
                }
                else
                {
                    lb_Rank04.Text = "";
                    lb_ID_04.Text = "";
                    lb_Name_04.Text = "";
                    lb_Score_04.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 3 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Rank=="-1")
                    {
                        lb_Rank04.Text = "";
                        lb_ID_04.Text = "";
                        lb_Name_04.Text = "";
                        lb_Score_04.Text = "";
                    }
                    else
                    {
                        lb_Rank04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].RealPlace;
                        lb_ID_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Number;
                        lb_Name_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Name;
                        lb_Score_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Rank;
                    }
                    
                }
                else
                {
                    lb_Rank04.Text = "";
                    lb_ID_04.Text = "";
                    lb_Name_04.Text = "";
                    lb_Score_04.Text = "";
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 5)
            {
                if (Now_Player - Display_Page * 6 - 4 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].RealPlace;
                    //lb_Rank05.ForeColor = Color.Black;
                    lb_ID_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Number;
                    //lb_ID_05.ForeColor = Color.Black;
                    lb_Name_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Name;
                    //lb_Name_05.ForeColor = Color.Black;
                    lb_Score_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Rank;
                    //lb_Score_05.ForeColor = Color.Black;
                    panel5.BackColor = Color.Red;
                    panel13.BackColor = Color.Yellow;
                    pnl_list05.BackgroundImage = comt.Properties.Resources.listred;
                }
                else
                {
                    lb_Rank05.Text = "";
                    lb_ID_05.Text = "";
                    lb_Name_05.Text = "";
                    lb_Score_05.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 4 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Rank=="-1")
                    {
                        lb_Rank05.Text = "";
                        lb_ID_05.Text = "";
                        lb_Name_05.Text = "";
                        lb_Score_05.Text = "";
                    }
                    else
                    {
                        lb_Rank05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].RealPlace;
                        lb_ID_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Number;
                        lb_Name_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Name;
                        lb_Score_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Rank;
                    }
                   
                }
                else
                {
                    lb_Rank05.Text = "";
                    lb_ID_05.Text = "";
                    lb_Name_05.Text = "";
                    lb_Score_05.Text = "";
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 0)
            {
                if (Now_Player - Display_Page * 6 - 5 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].RealPlace;
                    //lb_Rank06.ForeColor = Color.Black;
                    lb_ID_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Number;
                    //lb_ID_06.ForeColor = Color.Black;
                    lb_Name_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Name;
                    //lb_Name_06.ForeColor = Color.Black;
                    lb_Score_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Rank;
                    //lb_Score_06.ForeColor = Color.Black;
                    panel6.BackColor = Color.Red;
                    panel14.BackColor = Color.Yellow;
                    pnl_list06.BackgroundImage = comt.Properties.Resources.listred;
                }
                else
                {
                    lb_Rank06.Text = "";
                    lb_ID_06.Text = "";
                    lb_Name_06.Text = "";
                    lb_Score_06.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 5 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Rank =="-1")
                    {
                        lb_Rank06.Text = "";
                        lb_ID_06.Text = "";
                        lb_Name_06.Text = "";
                        lb_Score_06.Text = "";
                    }
                    else
                    {
                        lb_Rank06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].RealPlace;
                        lb_ID_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Number;
                        lb_Name_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Name;
                        lb_Score_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Rank;
                    }
                    
                }
                else
                {
                    lb_Rank06.Text = "";
                    lb_ID_06.Text = "";
                    lb_Name_06.Text = "";
                    lb_Score_06.Text = "";
                }
            }
            #endregion
            */

            #endregion

            #region mark

            /*
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 7)
            {
                if (Now_Player - Display_Page * 8 - 6 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].RealPlace;
                    lb_Rank07.ForeColor = Color.Black;
                    lb_ID_07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Number;
                    lb_ID_07.ForeColor = Color.Black;
                    lb_Name_07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Name;
                    lb_Name_07.ForeColor = Color.Black;
                    lb_Score_07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Rank;
                    lb_Score_07.ForeColor = Color.Black;
                    panel7.BackColor = Color.Red;
                    panel15.BackColor = Color.Yellow;
                }
                else
                {
                    lb_Rank07.Text = "";
                    lb_ID_07.Text = "";
                    lb_Name_07.Text = "";
                    lb_Score_07.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 8 - 6 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Rank=="-1")
                    {
                        lb_Rank07.Text = "";
                        lb_ID_07.Text = "";
                        lb_Name_07.Text = "";
                        lb_Score_07.Text = "";
                    }
                    else
                    {
                        lb_Rank07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].RealPlace;
                        lb_ID_07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Number;
                        lb_Name_07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Name;
                        lb_Score_07.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 6].Rank;
                    }
                   
                }
                else
                {
                    lb_Rank07.Text = "";
                    lb_ID_07.Text = "";
                    lb_Name_07.Text = "";
                    lb_Score_07.Text = "";
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 0)
            {
                if (Now_Player - Display_Page * 8 - 7 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].RealPlace;
                    lb_Rank08.ForeColor = Color.Black;
                    lb_ID_08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Number;
                    lb_ID_08.ForeColor = Color.Black;
                    lb_Name_08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Name;
                    lb_Name_08.ForeColor = Color.Black;
                    lb_Score_08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Rank;
                    lb_Score_08.ForeColor = Color.Black;
                    panel8.BackColor = Color.Red;
                    panel16.BackColor = Color.Yellow;
                }
                else
                {
                    lb_Rank08.Text = "";
                    lb_ID_08.Text = "";
                    lb_Name_08.Text = "";
                    lb_Score_08.Text = "";
                }
            }
            else
            {
                if (Now_Player - Display_Page * 8 - 7 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Rank=="-1")
                    {
                        lb_Rank08.Text = "";
                        lb_ID_08.Text = "";
                        lb_Name_08.Text = "";
                        lb_Score_08.Text = "";
                    }
                    else
                    {
                        lb_Rank08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].RealPlace;
                        lb_ID_08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Number;
                        lb_Name_08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Name;
                        lb_Score_08.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 7].Rank;
                    }
                   
                }
                else
                {
                    lb_Rank08.Text = "";
                    lb_ID_08.Text = "";
                    lb_Name_08.Text = "";
                    lb_Score_08.Text = "";
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            */
            #endregion

            //調整版面大小============================================================================================================
            #region 調整版面大小
            #region Name
            lb_Name_01.Font = new Font(lb_Name_01.Font.FontFamily, 60.0f, lb_Name_01.Font.Style);
            while (lb_Name_01.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name_01.Text,
            new Font(lb_Name_01.Font.FontFamily, lb_Name_01.Font.Size, lb_Name_01.Font.Style)).Width)
            {
                lb_Name_01.Font = new Font(lb_Name_01.Font.FontFamily, lb_Name_01.Font.Size - 0.5f, lb_Name_01.Font.Style);
            }

            lb_Name_02.Font = new Font(lb_Name_02.Font.FontFamily, 60.0f, lb_Name_02.Font.Style);
            while (lb_Name_02.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name_02.Text,
            new Font(lb_Name_02.Font.FontFamily, lb_Name_02.Font.Size, lb_Name_02.Font.Style)).Width)
            {
                lb_Name_02.Font = new Font(lb_Name_02.Font.FontFamily, lb_Name_02.Font.Size - 0.5f, lb_Name_02.Font.Style);
            }

            lb_Name_03.Font = new Font(lb_Name_03.Font.FontFamily, 60.0f, lb_Name_03.Font.Style);
            while (lb_Name_03.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name_03.Text,
            new Font(lb_Name_03.Font.FontFamily, lb_Name_03.Font.Size, lb_Name_03.Font.Style)).Width)
            {
                lb_Name_03.Font = new Font(lb_Name_03.Font.FontFamily, lb_Name_03.Font.Size - 0.5f, lb_Name_03.Font.Style);
            }

            lb_Name_04.Font = new Font(lb_Name_04.Font.FontFamily, 60.0f, lb_Name_04.Font.Style);
            while (lb_Name_04.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name_04.Text,
            new Font(lb_Name_04.Font.FontFamily, lb_Name_04.Font.Size, lb_Name_04.Font.Style)).Width)
            {
                lb_Name_04.Font = new Font(lb_Name_04.Font.FontFamily, lb_Name_04.Font.Size - 0.5f, lb_Name_04.Font.Style);
            }

            lb_Name_05.Font = new Font(lb_Name_05.Font.FontFamily, 60.0f, lb_Name_05.Font.Style);
            while (lb_Name_05.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name_05.Text,
            new Font(lb_Name_05.Font.FontFamily, lb_Name_05.Font.Size, lb_Name_05.Font.Style)).Width)
            {
                lb_Name_05.Font = new Font(lb_Name_05.Font.FontFamily, lb_Name_05.Font.Size - 0.5f, lb_Name_05.Font.Style);
            }

            lb_Name_06.Font = new Font(lb_Name_06.Font.FontFamily, 60.0f, lb_Name_06.Font.Style);
            while (lb_Name_06.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Name_06.Text,
            new Font(lb_Name_06.Font.FontFamily, lb_Name_06.Font.Size, lb_Name_06.Font.Style)).Width)
            {
                lb_Name_06.Font = new Font(lb_Name_06.Font.FontFamily, lb_Name_06.Font.Size - 0.5f, lb_Name_06.Font.Style);
            }
            #endregion
            //-----------------------------------------------------------------------------
            #region ID
            lb_ID_01.Font = new Font(lb_ID_01.Font.FontFamily, 60.0f, lb_ID_01.Font.Style);
            while (lb_ID_01.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID_01.Text,
            new Font(lb_ID_01.Font.FontFamily, lb_ID_01.Font.Size, lb_ID_01.Font.Style)).Width)
            {
                lb_ID_01.Font = new Font(lb_ID_01.Font.FontFamily, lb_ID_01.Font.Size - 0.5f, lb_ID_01.Font.Style);
            }

            lb_ID_02.Font = new Font(lb_ID_02.Font.FontFamily, 60.0f, lb_ID_02.Font.Style);
            while (lb_ID_02.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID_02.Text,
            new Font(lb_ID_02.Font.FontFamily, lb_ID_02.Font.Size, lb_ID_02.Font.Style)).Width)
            {
                lb_ID_02.Font = new Font(lb_ID_02.Font.FontFamily, lb_ID_02.Font.Size - 0.5f, lb_ID_02.Font.Style);
            }

            lb_ID_03.Font = new Font(lb_ID_03.Font.FontFamily, 60.0f, lb_ID_03.Font.Style);
            while (lb_ID_03.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID_03.Text,
            new Font(lb_ID_03.Font.FontFamily, lb_ID_03.Font.Size, lb_ID_03.Font.Style)).Width)
            {
                lb_ID_03.Font = new Font(lb_ID_03.Font.FontFamily, lb_ID_03.Font.Size - 0.5f, lb_ID_03.Font.Style);
            }

            lb_ID_04.Font = new Font(lb_ID_04.Font.FontFamily, 60.0f, lb_ID_04.Font.Style);
            while (lb_ID_04.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID_04.Text,
            new Font(lb_ID_04.Font.FontFamily, lb_ID_04.Font.Size, lb_ID_04.Font.Style)).Width)
            {
                lb_ID_04.Font = new Font(lb_ID_04.Font.FontFamily, lb_ID_04.Font.Size - 0.5f, lb_ID_04.Font.Style);
            }

            lb_ID_05.Font = new Font(lb_ID_05.Font.FontFamily, 60.0f, lb_ID_05.Font.Style);
            while (lb_ID_05.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID_05.Text,
            new Font(lb_ID_05.Font.FontFamily, lb_ID_05.Font.Size, lb_ID_05.Font.Style)).Width)
            {
                lb_ID_05.Font = new Font(lb_ID_05.Font.FontFamily, lb_ID_05.Font.Size - 0.5f, lb_ID_05.Font.Style);
            }

            lb_ID_06.Font = new Font(lb_ID_06.Font.FontFamily, 60.0f, lb_ID_06.Font.Style);
            while (lb_ID_06.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_ID_06.Text,
            new Font(lb_ID_06.Font.FontFamily, lb_ID_06.Font.Size, lb_ID_06.Font.Style)).Width)
            {
                lb_ID_06.Font = new Font(lb_ID_06.Font.FontFamily, lb_ID_06.Font.Size - 0.5f, lb_ID_06.Font.Style);
            }
            #endregion
            //-----------------------------------------------------------------------------
            #region Rank
            lb_Rank01.Font = new Font(lb_Rank01.Font.FontFamily, 60.0f, lb_Rank01.Font.Style);
            while (lb_Rank01.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Rank01.Text,
            new Font(lb_Rank01.Font.FontFamily, lb_Rank01.Font.Size, lb_Rank01.Font.Style)).Width)
            {
                lb_Rank01.Font = new Font(lb_Rank01.Font.FontFamily, lb_Rank01.Font.Size - 0.5f, lb_Rank01.Font.Style);
            }

            lb_Rank02.Font = new Font(lb_Rank02.Font.FontFamily, 60.0f, lb_Rank02.Font.Style);
            while (lb_Rank02.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Rank02.Text,
            new Font(lb_Rank02.Font.FontFamily, lb_Rank02.Font.Size, lb_Rank02.Font.Style)).Width)
            {
                lb_Rank02.Font = new Font(lb_Rank02.Font.FontFamily, lb_Rank02.Font.Size - 0.5f, lb_Rank02.Font.Style);
            }

            lb_Rank03.Font = new Font(lb_Rank03.Font.FontFamily, 60.0f, lb_Rank03.Font.Style);
            while (lb_Rank03.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Rank03.Text,
            new Font(lb_Rank03.Font.FontFamily, lb_Rank03.Font.Size, lb_Rank03.Font.Style)).Width)
            {
                lb_Rank03.Font = new Font(lb_Rank03.Font.FontFamily, lb_Rank03.Font.Size - 0.5f, lb_Rank03.Font.Style);
            }

            lb_Rank04.Font = new Font(lb_Rank04.Font.FontFamily, 60.0f, lb_Rank04.Font.Style);
            while (lb_Rank04.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Rank04.Text,
            new Font(lb_Rank04.Font.FontFamily, lb_Rank04.Font.Size, lb_Rank04.Font.Style)).Width)
            {
                lb_Rank04.Font = new Font(lb_Rank04.Font.FontFamily, lb_Rank04.Font.Size - 0.5f, lb_Rank04.Font.Style);
            }

            lb_Rank05.Font = new Font(lb_Rank05.Font.FontFamily, 60.0f, lb_Rank05.Font.Style);
            while (lb_Rank05.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Rank05.Text,
            new Font(lb_Rank05.Font.FontFamily, lb_Rank05.Font.Size, lb_Rank05.Font.Style)).Width)
            {
                lb_Rank05.Font = new Font(lb_Rank05.Font.FontFamily, lb_Rank05.Font.Size - 0.5f, lb_Rank05.Font.Style);
            }

            lb_Rank06.Font = new Font(lb_Rank06.Font.FontFamily, 60.0f, lb_Rank06.Font.Style);
            while (lb_Rank06.Width < System.Windows.Forms.TextRenderer.MeasureText(lb_Rank06.Text,
            new Font(lb_Rank06.Font.FontFamily, lb_Rank06.Font.Size, lb_Rank06.Font.Style)).Width)
            {
                lb_Rank06.Font = new Font(lb_Rank06.Font.FontFamily, lb_Rank06.Font.Size - 0.5f, lb_Rank06.Font.Style);
            }
            #endregion
            #endregion
            //調整版面大小============================================================================================================
            //IsWritng_Rank確認狀態/////////////////////////////////
            
            comt.Form1.WriteExcelAll();
            comt.Form1.IsWritng_Rank = false;//寫入完畢，寫入完畢後才開放Form1可以轉跳至下一位選手
            //timer_bling.Enabled = true;
            //timer_flash.Enabled = true;
            //IsWritng_Rank確認狀態/////////////////////////////////
            //顯示/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            /*
            if (Select_Pose == "1")
            {

            }
            if (Select_Pose == "2")
            { 
            
            }
            */
        }

        //ASCII  Hex  互轉////////////////////////////////////////////////////////////////////
        #region ASCII  Hex  互轉
        public static int ASC(string S)
        {
            int N = Convert.ToInt32(S[0]);
            return N;
        }
        public static char Chr(int Num)
        {
            char C = Convert.ToChar(Num);
            return C;
        }

        /*
        public static int ASC(char C)
        {
            int N = Convert.ToInt32(C);
            return N;
        }
        */
        #endregion
        //ASCII  Hex  互轉////////////////////////////////////////////////////////////////////

        //數字轉國字///////////////////////////////////////////////
        #region 數字轉國字
        static string GetChineseNumber(int number)
        {
            string[] chineseNumber = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string[] unit = { "", "十", "百", "千", "萬" };
            string ret = "";
            string inputNumber = number.ToString();
            int i = inputNumber.Length - 1;
            bool b = false;
            foreach (char c in inputNumber)
            {
                if (c > '0')
                {
                    if (b)
                    {
                        ret += chineseNumber[0];
                        b = false;
                    }
                    ret += chineseNumber[(int)(c - '0')] + unit[i];
                }
                else
                    b = true;
                i--;
            }
            /////////////////////////////////////////
            if (ret.Length >= 2)
            {
                if (ret.Substring(0, 2) == "一十")
                {
                    ret = ret.Replace("一十", "十");
                }
            }
            /////////////////////////////////////////
            return ret;
        }
        #endregion
        //數字轉國字///////////////////////////////////////////////

        private void Rank_FormClosing(object sender, FormClosingEventArgs e)
        {
            comt.Form1.IsOpen_Rank = false;//因為已經開啟Demonstrate的視窗了，所以將Rank關掉後將IsOpen_Rank設為false
            comt.Form1.IsOpen_Rank_ByRating = false;//因為已經開啟Demonstrate的視窗了，所以將Rank關掉後將IsOpen_Rank設為false
        }

        private void timer_Rank_Tick(object sender, EventArgs e)
        {

        }

        private void Rank_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)//按下ESC  
            {
                this.Close();
            }
        }
        private void Rank_Load(object sender, EventArgs e)
        {

        }

        int Count_Times_Bling = 0;
        int times = 10;//閃爍次數
        bool x = true;
        private void timer_bling_Tick(object sender, EventArgs e)
        {
            #region mark
            /*
            Count_Times_Bling++;
            if (Count_Times_Bling == (2 * times + 1))
            {
                pb_bg_bling.BackColor = Color.Transparent;//讓他閃到最後變透明，就是沒有框
                Count_Times_Bling = 0;
                timer_bling.Enabled = false;
            }
            else 
            {
                if (x == true)
                {
                    pb_bg_bling.BackColor = Color.Red;
                    x = false;
                }
                else
                {
                    pb_bg_bling.BackColor = Color.Transparent;
                    x = true;
                }
            }
            */            
            //Rank = new Rank();
            //Rank.Activate();        
            #endregion

            //Sam/////////////////////////////////////////////
            Graphics g = this.CreateGraphics();
            drawRank _drawRank = new drawRank();
            //_drawRank.drawEmptyString(g, 47, 339, "22", "22", "王大明", "10.00");
            /*
            _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, "22", "22", "王大明", "10.00");
            _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, "22", "22", "王大明", "10.00");
            _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, "22", "22", "王大明", "10.00");
            _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, "22", "22", "王大明", "10.00");
            _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, "22", "22", "王大明", "10.00");
            _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, "22", "22", "王大明", "10.00");
            

            _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List1, "11", "11", "11", "11");//sam
            */
            //Sam/////////////////////////////////////////////
            /*
            //Location///////////////////////////////////////////////////
            int Location_Bar_X = 44;
            int Location_Bar_Y_List1= 139;
            int Location_Bar_Y_List2 = 243;
            int Location_Bar_Y_List3 = 350;
            int Location_Bar_Y_List4 = 455;
            int Location_Bar_Y_List5 = 562;
            int Location_Bar_Y_List6 = 662;
            //Location///////////////////////////////////////////////////
            //前後都補0的數字字串 
            Console.WriteLine(String.Format("{0:0000.0000}", 157.42)); // 輸出 0157.4200
            String.Format("{0:0.0000}", 157.42)
            //金額的表示_改良_取到小數2位 
            Console.WriteLine(String.Format("{0:$###,###,###,##0.00}", 0)); // $0.00 
            Console.WriteLine(String.Format("{0:$###,###,###,##0.00}", 12.5)); // $12.50 
            Console.WriteLine(String.Format("{0:$###,###,###,##0.00}", 3456234532)); // $3,456,234,532.0
            String.Format("{0:#0.00}", )
            */

            #region Display
            ////找出現在選手的名次///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            int Now_Player_Place = 1;//找出現在選手的名次
            for (int m = Real_AlreadyRatedPlayerCount; m >= 1; m--)////Update_20140518_15//for (int m = Now_Player; m >= 1; m--)
            {
                if (comt.Form1.ImportData[2, Now_Player] == ScoreRank_Total_Variable[m].Name && comt.Form1.ImportData[15, Now_Player] == ScoreRank_Total_Variable[m].Number)
                {
                    Now_Player_Place = Real_AlreadyRatedPlayerCount - m + 1;//Update_20140518_16//Now_Player_Place = Now_Player - m + 1
                }
            }
            ////找出現在選手的名次///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            //顯示/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            int Display_Place;//顯示名次的位置
            int Display_Page;//顯示名次的頁碼
            Display_Place = Now_Player_Place % 6;
            if (Display_Place == 0)
            {
                Display_Page = (Now_Player_Place / 6) - 1;
            }
            else
            {
                Display_Page = Now_Player_Place / 6;
            }
            Now_Player = Real_AlreadyRatedPlayerCount;////Update_20140518_17//本來沒有

            /*
            string k1 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8].Number;
            string k2 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8].Name;
            string k2_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8].Rank;
            string k3 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 1].Number;
            string k4 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 1].Name;
            string k4_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 1].Rank;
            string k11 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8-2].Number;
            string k21 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8-2].Name;
            string k21_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8-2].Rank;
            string k31 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 3].Number;
            string k41 = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 3].Name;
            string k41_ = ScoreRank_Total_Variable[Now_Player - Display_Page * 8 - 3].Rank;
            */
            /*
            for (int i = 0; i <= Now_Player; i++)
            {
                textBox1.AppendText(ScoreRank_Total_Variable[i].Name + "  " +ScoreRank_Total_Variable[i].Score + "\n");
            }
            */

            if (Display_Place == 1)
            {
                if (Now_Player - Display_Page * 6 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].RealPlace; //(1 + Display_Page * 8).ToString();
                    //lb_Rank01.ForeColor = Color.Black;
                    lb_ID_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Number;
                    //lb_ID_01.ForeColor = Color.Black;
                    lb_Name_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Name;
                    //lb_Name_01.ForeColor = Color.Black;
                    lb_Score_01.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Rank));
                    //lb_Score_01.ForeColor = Color.Black;
                    panel1.BackColor = Color.Red;
                    panel9.BackColor = Color.Yellow;
                    pnl_list01.BackgroundImage = comt.Properties.Resources.listred;
                    _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                }
                else
                {
                    lb_Rank01.Text = "";
                    lb_ID_01.Text = "";
                    lb_Name_01.Text = "";
                    lb_Score_01.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, "", "", "", "");//sam
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Rank == "-1")
                    {
                        lb_Rank01.Text = "";
                        lb_ID_01.Text = "";
                        lb_Name_01.Text = "";
                        lb_Score_01.Text = "";
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, "", "", "", "");//sam
                    }
                    else
                    {
                        lb_Rank01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].RealPlace;
                        lb_ID_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Number;
                        lb_Name_01.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Name;
                        lb_Score_01.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6].Rank));
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                    }
                }
                else
                {
                    lb_Rank01.Text = "";
                    lb_ID_01.Text = "";
                    lb_Name_01.Text = "";
                    lb_Score_01.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, "", "", "", "");//sam
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 2)
            {
                if (Now_Player - Display_Page * 6 - 1 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].RealPlace;
                    //lb_Rank02.ForeColor = Color.Black;
                    lb_ID_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Number;
                    //lb_ID_02.ForeColor = Color.Black;
                    lb_Name_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Name;
                    //lb_Name_02.ForeColor = Color.Black;
                    lb_Score_02.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Rank));
                    //lb_Score_02.ForeColor = Color.Black;
                    panel2.BackColor = Color.Red;
                    panel10.BackColor = Color.Yellow;
                    pnl_list02.BackgroundImage = comt.Properties.Resources.listred;
                    _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                }
                else
                {
                    lb_Rank02.Text = "";
                    lb_ID_02.Text = "";
                    lb_Name_02.Text = "";
                    lb_Score_02.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, "", "", "", "");//sam
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 1 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Rank == "-1")
                    {
                        lb_Rank02.Text = "";
                        lb_ID_02.Text = "";
                        lb_Name_02.Text = "";
                        lb_Score_02.Text = "";
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, "", "", "", "");//sam
                    }
                    else
                    {
                        lb_Rank02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].RealPlace;
                        lb_ID_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Number;
                        lb_Name_02.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Name;
                        lb_Score_02.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 1].Rank));
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                    }

                }
                else
                {
                    lb_Rank02.Text = "";
                    lb_ID_02.Text = "";
                    lb_Name_02.Text = "";
                    lb_Score_02.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, "", "", "", "");//sam
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 3)
            {
                if (Now_Player - Display_Page * 6 - 2 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].RealPlace;
                    //lb_Rank03.ForeColor = Color.Black;
                    lb_ID_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Number;
                    //lb_ID_03.ForeColor = Color.Black;
                    lb_Name_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Name;
                    //lb_Name_03.ForeColor = Color.Black;
                    lb_Score_03.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Rank));
                    //lb_Score_03.ForeColor = Color.Black;
                    panel3.BackColor = Color.Red;
                    panel11.BackColor = Color.Yellow;
                    pnl_list03.BackgroundImage = comt.Properties.Resources.listred;
                    _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                }
                else
                {
                    lb_Rank03.Text = "";
                    lb_ID_03.Text = "";
                    lb_Name_03.Text = "";
                    lb_Score_03.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, "", "", "", "");//sam
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 2 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Rank == "-1")
                    {
                        lb_Rank03.Text = "";
                        lb_ID_03.Text = "";
                        lb_Name_03.Text = "";
                        lb_Score_03.Text = "";
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, "", "", "", "");//sam
                    }
                    else
                    {
                        lb_Rank03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].RealPlace;
                        lb_ID_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Number;
                        lb_Name_03.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Name;
                        lb_Score_03.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 2].Rank));
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                    }

                }
                else
                {
                    lb_Rank03.Text = "";
                    lb_ID_03.Text = "";
                    lb_Name_03.Text = "";
                    lb_Score_03.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, "", "", "", "");//sam
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 4)
            {
                if (Now_Player - Display_Page * 6 - 3 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].RealPlace;
                    //lb_Rank04.ForeColor = Color.Black;
                    lb_ID_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Number;
                    //lb_ID_04.ForeColor = Color.Black;
                    lb_Name_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Name;
                    //lb_Name_04.ForeColor = Color.Black;
                    lb_Score_04.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Rank));
                    //lb_Score_04.ForeColor = Color.Black;
                    panel4.BackColor = Color.Red;
                    panel12.BackColor = Color.Yellow;
                    pnl_list04.BackgroundImage = comt.Properties.Resources.listred;
                    _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                }
                else
                {
                    lb_Rank04.Text = "";
                    lb_ID_04.Text = "";
                    lb_Name_04.Text = "";
                    lb_Score_04.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, "", "", "", "");//sam
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 3 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Rank == "-1")
                    {
                        lb_Rank04.Text = "";
                        lb_ID_04.Text = "";
                        lb_Name_04.Text = "";
                        lb_Score_04.Text = "";
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, "", "", "", "");//sam
                    }
                    else
                    {
                        lb_Rank04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].RealPlace;
                        lb_ID_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Number;
                        lb_Name_04.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Name;
                        lb_Score_04.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 3].Rank));
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                    }

                }
                else
                {
                    lb_Rank04.Text = "";
                    lb_ID_04.Text = "";
                    lb_Name_04.Text = "";
                    lb_Score_04.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, "", "", "", "");//sam
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 5)
            {
                if (Now_Player - Display_Page * 6 - 4 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].RealPlace;
                    //lb_Rank05.ForeColor = Color.Black;
                    lb_ID_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Number;
                    //lb_ID_05.ForeColor = Color.Black;
                    lb_Name_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Name;
                    //lb_Name_05.ForeColor = Color.Black;
                    lb_Score_05.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Rank));
                    //lb_Score_05.ForeColor = Color.Black;
                    panel5.BackColor = Color.Red;
                    panel13.BackColor = Color.Yellow;
                    pnl_list05.BackgroundImage = comt.Properties.Resources.listred;
                    _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                }
                else
                {
                    lb_Rank05.Text = "";
                    lb_ID_05.Text = "";
                    lb_Name_05.Text = "";
                    lb_Score_05.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, "", "", "", "");//sam
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 4 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Rank == "-1")
                    {
                        lb_Rank05.Text = "";
                        lb_ID_05.Text = "";
                        lb_Name_05.Text = "";
                        lb_Score_05.Text = "";
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, "", "", "", "");//sam
                    }
                    else
                    {
                        lb_Rank05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].RealPlace;
                        lb_ID_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Number;
                        lb_Name_05.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Name;
                        lb_Score_05.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 4].Rank));
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                    }

                }
                else
                {
                    lb_Rank05.Text = "";
                    lb_ID_05.Text = "";
                    lb_Name_05.Text = "";
                    lb_Score_05.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, "", "", "", "");//sam
                }
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            if (Display_Place == 0)
            {
                if (Now_Player - Display_Page * 6 - 5 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    lb_Rank06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].RealPlace;
                    //lb_Rank06.ForeColor = Color.Black;
                    lb_ID_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Number;
                    //lb_ID_06.ForeColor = Color.Black;
                    lb_Name_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Name;
                    //lb_Name_06.ForeColor = Color.Black;
                    lb_Score_06.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Rank));
                    //lb_Score_06.ForeColor = Color.Black;
                    panel6.BackColor = Color.Red;
                    panel14.BackColor = Color.Yellow;
                    pnl_list06.BackgroundImage = comt.Properties.Resources.listred;
                    _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                }
                else
                {
                    lb_Rank06.Text = "";
                    lb_ID_06.Text = "";
                    lb_Name_06.Text = "";
                    lb_Score_06.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, "", "", "", "");//sam
                }
            }
            else
            {
                if (Now_Player - Display_Page * 6 - 5 >= 1)//大於等於一 才是 Rank後的 ScoreRank_Total_Variable陣列中分數最小的 ， 也就是 位置0 的 (score欄位)數值0 是強塞進去的最小值 不予計算
                {
                    if (ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Rank == "-1")
                    {
                        lb_Rank06.Text = "";
                        lb_ID_06.Text = "";
                        lb_Name_06.Text = "";
                        lb_Score_06.Text = "";
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, "", "", "", "");//sam
                    }
                    else
                    {
                        lb_Rank06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].RealPlace;
                        lb_ID_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Number;
                        lb_Name_06.Text = ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Name;
                        lb_Score_06.Text = String.Format("{0:#0.00}", float.Parse(ScoreRank_Total_Variable[Now_Player - Display_Page * 6 - 5].Rank));
                        _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                    }

                }
                else
                {
                    lb_Rank06.Text = "";
                    lb_ID_06.Text = "";
                    lb_Name_06.Text = "";
                    lb_Score_06.Text = "";
                    _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, "", "", "", "");//sam
                }
            }
            #endregion


            int bling_interval = 800;
            if (Display_Place == 1)
            {
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List1, lb_Rank01.Text, lb_ID_01.Text, lb_Name_01.Text, lb_Score_01.Text);//sam
            }
            else if (Display_Place == 2)
            {
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List2, lb_Rank02.Text, lb_ID_02.Text, lb_Name_02.Text, lb_Score_02.Text);//sam
            }
            else if (Display_Place == 3)
            {
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List3, lb_Rank03.Text, lb_ID_03.Text, lb_Name_03.Text, lb_Score_03.Text);//sam
            }
            else if (Display_Place == 4)
            {
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List4, lb_Rank04.Text, lb_ID_04.Text, lb_Name_04.Text, lb_Score_04.Text);//sam
            }
            else if (Display_Place == 5)
            {
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List5, lb_Rank05.Text, lb_ID_05.Text, lb_Name_05.Text, lb_Score_05.Text);//sam
            }
            else if (Display_Place == 0)
            {
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawBlueString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
                System.Threading.Thread.Sleep(bling_interval);
                _drawRank.drawRedString(g, Location_Bar_X, Location_Bar_Y_List6, lb_Rank06.Text, lb_ID_06.Text, lb_Name_06.Text, lb_Score_06.Text);//sam
            }

            timer_bling.Enabled = false;
        }

        private void timer_flash_Tick(object sender, EventArgs e)
        {
            if (x == true)
            {
                pb_bg_bling.BackColor = Color.Red;
                x = false;
            }
            else
            {
                pb_bg_bling.BackColor = Color.Transparent;
                x = true;
            }
        }


      
    }
}
