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
using System.Threading;
using System.Web.UI;

//最後改版
//20130914:加入紀錄伺服器IP功能及上傳時詢問功能(ie,1.先詢問是否匯出(選擇匯出則Excel為有排序，選否則不排序) 2.再詢問是否上傳，若失敗則繼續停留在上傳畫面 PS:並於每步驟後詢問使用者是否關閉程式)
//20130822:FTP上傳功能
//20130818:原始分數加入
//20130729:改版內容:加入閃爍，修正重新評分。
//20130730:將result頁面總分顯示公式改成和控制端總分公式相同。
//20130808:將excel表格的import及writeexcelall部分整理完成
//20130519_2316
//20131215:找出Bug在Rank內的公式計算變數Final_Weighted_Score
//20140518:1.在Form1加入匯出時會再次的排序的函式 RankForAllPlayer()
//         2.在Rank加入Update_20140518，若回來重評棄權的人時，才可以知道之前評到第幾個人，再一起做排序/////////////////////
//
/*
//移動視窗到延伸桌面
this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
this.DesktopLocation = new System.Drawing.Point(2000, 0);
*/
namespace comt
{
    public partial class Form1 : Form
    {
        //變數///////////////////////////////////////////////////////////
        string Header_Byte1, Header_Byte2;
        //-----------------------------------
        string Device_Byte1, Device_Byte2;
        //-----------------------------------
        string Opcode_Byte1, Opcode_Byte2;
        //-----------------------------------
        string Data1_Byte1, Data1_Byte2;
        string Data2_Byte1, Data2_Byte2;
        string Data3_Byte1, Data3_Byte2;
        //-----------------------------------
        string CRC_Byte1, CRC_Byte2, CRC_Byte3, CRC_Byte4;
        //----------------------------------------------------------
        string Receive_Instruction;
        //----------------------------------------------------------
        string DeviceName;

        string CRC_StartRating;
        string CRC_EndRating;
        string CRC_PCAskForData;
        string CRC_RoutineAck;
  
        byte[] Byte12_StartRating;
        byte[] Byte12_EndRating;
        byte[] Byte12_PCAskForData;
        byte[] Byte12_RoutineAck;

        byte[] Byte16_StartRating;
        byte[] Byte16_EndRating;
        byte[] Byte16_PCAskForData;
        byte[] Byte16_RoutineAck;
        //變數///////////////////////////////////////////////////////////

        //選手資料///////////////////////////////////////////////////////
        public static string[,] ImportData;        
        public static string Select_Pose;//共有多少品勢
        public static int Raw_CompetitorNumber;//有多少位參賽者
        public static int Column_Pose1;//單品勢有多少欄位
        public static int Column_Pose2;//雙品勢有多少欄位
        public static string[,] ImportData_Score;    
        //--------------------------------        
        public static int Count_Referee;//看有幾個評審
        //選手資料///////////////////////////////////////////////////////

        //現行狀態State//////////////////////////////////////////////////
        public static int Now_Player = 1;
        public static int Now_Pose = 1;
        public static int Now_Device = 1;
        public static int Now_Wait_Device = 1;
        public static int Milestone_Now_Device = 1;//以現在這個機台為里程碑，下次又找到他代表全部找完(ie,全部數值都蒐集完成)

        public static string[] Flag_Correctness;//看正確性評分了沒
        public static string[] Flag_Performance01;//看表現性性評分了沒
        public static string[] Flag_Performance02;//看表現性性評分了沒
        public static string[] Flag_Performance03;//看表現性性評分了沒
        public static string[] Score_Correctness;//填入正確性分數
        public static string[] Score_Performance01;//填入表現性分數
        public static string[] Score_Performance02;//填入表現性分數
        public static string[] Score_Performance03;//填入表現性分數


        public static int[] Flag_Pooling;//用來檢查是否有Ack//若5次以上沒Ack則表示Device無回應        
        //現行狀態State//////////////////////////////////////////////////

        //mutex//////////////////////////////////////////////////////////
        bool Mutex_DeviceRoutineAck = true;
        //mutex//////////////////////////////////////////////////////////

        //視窗///////////////////////////////////////////////////////////
        Demonstrate Demonstrate;
        Rating Rating;
        public static bool IsOpen_Demonstrate = false;
        bool IsOpen_Rating = false;
        public static bool IsOpen_Rank = false;
        bool IsOpen_ImportSetting = false;        
        //視窗///////////////////////////////////////////////////////////

        //Excel檔案位置//////////////////////////////////////////////////////
        public static string ExcelFileLocation;
        //Excel檔案位置//////////////////////////////////////////////////////

        //RemoteControl_Timer(開啟Demonstrate的Timer)////////////////////////////////////
        public static bool RemoteControl_TimerStart = false;
        //RemoteControl_Timer(開啟Demonstrate的Timer)////////////////////////////////////

        //視窗//////////////////////////////
        public static bool IsOpen_Rank_ByRating = false;
        //視窗//////////////////////////////

        //ExcelAccessSecurityState///////////////////////////////////////////////////////
        public static bool IsWritng_Rank = false;//去看Rank頁面的Excel寫入(ie,ExcelWriteAll是否已經完整寫入)
        //ExcelAccessSecurityState///////////////////////////////////////////////////////

        string msg = "";


        //版本2 Declaration
        int[] DeviceState = new int[7];
        int[] DisconnectCnt = new int[7];
        int[] RefereeStateDouble = new int[7];
        double[] RefereeAcuraccyScore = new double[7];
        double[] RefereePresentionScore = new double[7];
        bool[] isFinished;
        public static float[][] ScoreArray = new float[Constant.ScoringItemNum][];
        int Referee = 2;
        System.Windows.Forms.Label[] Lb_Device;
        public Form1()
        {
            InitializeComponent();
            //timer_Scoring.Stop();
            /*
            //移動視窗到延伸桌面
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            */

            cb_COM.Text = God.Yuan.GetCOMPorts()[0];
            Lb_Device = new System.Windows.Forms.Label[] { lb_device01, lb_device02, lb_device03, lb_device04, lb_device05, lb_device06, lb_device07};
            
            isFinished = new bool[] { false, false, false, false };
            RefereeStateDouble = new int[7] {  0, 0, 0, 0, 0, 0, 0  };
            timer_Scoring.Stop();
          
            #region mark01
            //label2.Text = ASC("A").ToString();
            //!;010A405060
            /*
            msg += Convert.ToString(ASC("!"), 16) + "\n"
                         + Convert.ToString(ASC(";"), 16) + "\n"
                         + Convert.ToString(ASC("0"), 16) + "\n"
                         + Convert.ToString(ASC("1"), 16) + "\n"
                         + Convert.ToString(ASC("0"), 16) + "\n"
                         + Convert.ToString(ASC("A"), 16) + "\n"
                         + Convert.ToString(ASC("4"), 16) + "\n"
                         + Convert.ToString(ASC("0"), 16) + "\n"
                         + Convert.ToString(ASC("5"), 16) + "\n"
                         + Convert.ToString(ASC("0"), 16) + "\n"
                         + Convert.ToString(ASC("6"), 16) + "\n"
                         + Convert.ToString(ASC("0"), 16) + Environment.NewLine;
            */

            //21 3b 30 31 30 41 34 30 35 30 36 30
            //48:65:61:72:74:20:52:61:74:65:20:53:65:6E:73:6F:72
            //int h = Convert.ToInt32(Chr(Convert.ToInt32("39", 16)).ToString()) + 19;
            /*
            msg += Chr(Convert.ToInt32("48", 16)) + "\n"
                          + Chr(Convert.ToInt32("39", 16)) + "\n"
                          + Chr(Convert.ToInt32("61", 16)) + "\n"
                          + Chr(Convert.ToInt32("72", 16)) + "\n"
                          + Chr(Convert.ToInt32("74", 16)) + "\n"
                          + Chr(Convert.ToInt32("20", 16)) + "\n"
                          + Chr(Convert.ToInt32("52", 16)) + "\n"
                          + Chr(Convert.ToInt32("61", 16)) + "\n"
                          + Chr(Convert.ToInt32("74", 16)) + "\n"
                          + Chr(Convert.ToInt32("65", 16)) + "\n"
                          + Chr(Convert.ToInt32("20", 16)) + "\n"
                          + Chr(Convert.ToInt32("53", 16)) + "\n"
                          + Chr(Convert.ToInt32("65", 16)) + "\n"
                          + Chr(Convert.ToInt32("6E", 16)) + "\n"
                          + Chr(Convert.ToInt32("73", 16)) + "\n"
                          + Chr(Convert.ToInt32("6F", 16)) + "\n"
                          + Chr(Convert.ToInt32("72", 16)) + "\n";
            textBox1.Text = msg;
            */
            #endregion            
        }
        string Real_Instruction;
        int firstCharacter;


        //字串擷取//////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static string SubString(string strDD, int startIndex, int length) //解決字串擷取超過範圍的問題(新類別）
        {
            int intLen = strDD.Length;//輸入的字串總共長度 18
            int intSubLen = intLen - startIndex;//算剩下的長度夠不夠你擷取"55 66 21 3B 30 31 " 18 - 12
            string strReturn;

            if (length == 0)
                strReturn = "";
            else
            {
                if (intLen <= startIndex)
                    strReturn = "";
                else
                {
                    if (length > intSubLen)
                        length = intSubLen;

                    strReturn = strDD.Substring(startIndex, length);
                }
            }
            return strReturn;
        }
        //字串擷取//////////////////////////////////////////////////////////////////////////////////////////////////////////////

        string string_sum = "";
        /*private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            int ByteSize = serialPort1.BytesToRead;//read一次時收到的data buffer size
            byte[] BufferData = new byte[ByteSize];
            //////////////////////////////////////////////////////////
            int bytes = serialPort1.BytesToRead;
            // create a byte array to hold the awaiting data
            byte[] comBuffer = new byte[bytes];
            // read the data and store it
            serialPort1.Read(comBuffer, 0, bytes);
            //display the data to the user
            //DisplayData(MessageType.Incoming, ByteToHex(comBuffer) + "\n");
            string_sum = string_sum + ByteToHex(comBuffer);
            firstCharacter = string_sum.IndexOf("21 3B ");
            try
            {
                if (firstCharacter != -1 && string_sum.Length >= firstCharacter + 48)
                {
                    while (string_sum.Length >= 48)
                    {
                        Real_Instruction = string_sum.Substring(firstCharacter, 48);//集合成一指令
                        string_sum = SubString(string_sum, firstCharacter + 48, (string_sum.Length - firstCharacter - 48));//扣除掉該指令，剩下的指令

                        if (Real_Instruction.Length == 48 && Real_Instruction.Substring(0, 5) == "21 3B")//先檢查資料長度有無到&&標頭檔對不對
                        {
                            #region if內
                            byte[] Command_Byte_Array = new byte[12] 
                            { 
                                Convert.ToByte("0x"+Real_Instruction.Substring(0,2), 16),//1
                                Convert.ToByte("0x"+Real_Instruction.Substring(3,2), 16),//2
                                Convert.ToByte("0x"+Real_Instruction.Substring(6,2), 16),//3
                                Convert.ToByte("0x"+Real_Instruction.Substring(9,2), 16),//4
                                Convert.ToByte("0x"+Real_Instruction.Substring(12,2), 16),//5
                                Convert.ToByte("0x"+Real_Instruction.Substring(15,2), 16),//6
                                Convert.ToByte("0x"+Real_Instruction.Substring(18,2), 16),//7
                                Convert.ToByte("0x"+Real_Instruction.Substring(21,2), 16),//8
                                Convert.ToByte("0x"+Real_Instruction.Substring(24,2), 16),//9
                                Convert.ToByte("0x"+Real_Instruction.Substring(27,2), 16),//10
                                Convert.ToByte("0x"+Real_Instruction.Substring(30,2), 16),//11
                                Convert.ToByte("0x"+Real_Instruction.Substring(33,2), 16),//12
                            };

                            string String_CRC16;
                            String_CRC16 = String.Format("{0,4:X}", CRC_any(Command_Byte_Array, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
                            String_CRC16 = String_CRC16.Replace(" ", "0");//不可少CRC若算出"0A58"可補0

                            if (Real_Instruction.Substring(42, 2) + Real_Instruction.Substring(45, 2) == String_CRC16)//再檢查CRC對不對
                            {
                                this.textBox1.Invoke(
                                    new MethodInvoker(
                                    delegate
                                    {
                                        this.textBox1.AppendText(Environment.NewLine);
                                        this.textBox1.AppendText("RX : " + Real_Instruction);
                                        Receive_Instruction = Real_Instruction;
                                        Real_Instruction = "";
                                        //this.textBox1.Text += " ";
                                    }
                                )
                                );

                                Header_Byte1 = Receive_Instruction.Substring(0, 2);
                                Header_Byte2 = Receive_Instruction.Substring(3, 2);
                                //-----------------------------------
                                Device_Byte1 = Receive_Instruction.Substring(6, 2);
                                Device_Byte2 = Receive_Instruction.Substring(9, 2);
                                //-----------------------------------
                                Opcode_Byte1 = Receive_Instruction.Substring(12, 2);
                                Opcode_Byte2 = Receive_Instruction.Substring(15, 2);
                                //-----------------------------------
                                Data1_Byte1 = Receive_Instruction.Substring(18, 2);
                                Data1_Byte2 = Receive_Instruction.Substring(21, 2);
                                Data2_Byte1 = Receive_Instruction.Substring(24, 2);
                                Data2_Byte2 = Receive_Instruction.Substring(27, 2);
                                Data3_Byte1 = Receive_Instruction.Substring(30, 2);
                                Data3_Byte2 = Receive_Instruction.Substring(33, 2);
                                //-----------------------------------
                                CRC_Byte1 = Receive_Instruction.Substring(36, 2);
                                CRC_Byte2 = Receive_Instruction.Substring(39, 2);
                                CRC_Byte3 = Receive_Instruction.Substring(42, 2);
                                CRC_Byte4 = Receive_Instruction.Substring(45, 2);

                                string Receive_OPcode = (Chr(Convert.ToInt32(Opcode_Byte1, 16)).ToString() + Chr(Convert.ToInt32(Opcode_Byte2, 16)).ToString());
                                int Receive_DeviceID = Convert.ToInt32(Chr(Convert.ToInt32(Device_Byte2, 16)).ToString());

                                if (Receive_OPcode == "02")//如果尚未輸入任何數值，只是傳回RoutuneAck
                                {
                                    if (Receive_DeviceID.ToString() == Now_Device.ToString())//確定了機碼，收回Ack
                                    {
                                        Flag_Pooling[Receive_DeviceID] = 0;//有收到Ack,就把它歸0繼續往下一台機器
                                        Milestone_Now_Device = Now_Device;//設里程碑，若找過一輪都有收到數值了則結束

                                        if (Now_Wait_Device == Receive_DeviceID)
                                        {
                                            //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                            #region 決定下一個Device要去polling誰
                                            if (Count_Referee == 1)//如果只有1台Device
                                            {
                                                Now_Device = 1;//則接下來Polling的Deviece都是1
                                            }
                                            else //若有1個以上的Device
                                            {
                                                Now_Device++;//往下一個Device前進
                                                if (Now_Device == Count_Referee)//由於NowDevice若和機台個數相等mod會變成0
                                                {
                                                    Now_Device = Now_Device;//所以若相等則讓他變成家王以後的機號而不做mod
                                                }
                                                else
                                                {
                                                    Now_Device = (Now_Device % Count_Referee);//若加完不為機台個數則去mod
                                                }
                                            }

                                            while (Flag_Correctness[Now_Device] == "1" && Flag_Performance01[Now_Device] == "1" && Flag_Performance02[Now_Device] == "1" && Flag_Performance03[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
                                            {
                                                if (Count_Referee == 1)
                                                {
                                                    Now_Device = 1;
                                                }
                                                else
                                                {
                                                    Now_Device++;
                                                    if (Now_Device == Count_Referee)
                                                    {
                                                        Now_Device = Now_Device;
                                                    }
                                                    else
                                                    {
                                                        Now_Device = (Now_Device % Count_Referee);
                                                    }
                                                }

                                                if (Flag_Correctness[Now_Device] == "0" || Flag_Performance01[Now_Device] == "0" || Flag_Performance02[Now_Device] == "0" || Flag_Performance03[Now_Device] == "0")
                                                {
                                                    break;
                                                }
                                                if (Milestone_Now_Device == Now_Device)//代表全部找完
                                                {
       
                                                    break;
                                                }
                                            }
                                            #endregion
                                            //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                        }

                                    }
                                }
                                if (Receive_OPcode == "03")
                                {
                                    Flag_Pooling[Receive_DeviceID] = 0;//有收到Ack,就把它歸0繼續往下一台機器
                                    Flag_Correctness[Receive_DeviceID] = "1";
                                    Score_Correctness[Receive_DeviceID] = Chr(Convert.ToInt32(Data1_Byte1, 16)).ToString() + "." + Chr(Convert.ToInt32(Data1_Byte2, 16)).ToString();

                                    Milestone_Now_Device = Now_Device;//設里程碑，若找過一輪都有收到數值了則結束

                                    if (Now_Wait_Device == Receive_DeviceID)
                                    {
                                        //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                        #region 決定下一個Device要去polling誰
                                        if (Count_Referee == 1)//如果只有1台Device
                                        {
                                            Now_Device = 1;//則接下來Polling的Deviece都是1
                                        }
                                        else //若有1個以上的Device
                                        {
                                            Now_Device++;//往下一個Device前進
                                            if (Now_Device == Count_Referee)//由於NowDevice若和機台個數相等mod會變成0
                                            {
                                                Now_Device = Now_Device;//所以若相等則讓他變成家王以後的機號而不做mod
                                            }
                                            else
                                            {
                                                Now_Device = (Now_Device % Count_Referee);//若加完不為機台個數則去mod
                                            }
                                        }

                                        while (Flag_Correctness[Now_Device] == "1" && Flag_Performance01[Now_Device] == "1" && Flag_Performance02[Now_Device] == "1" && Flag_Performance03[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
                                        {
                                            if (Count_Referee == 1)
                                            {
                                                Now_Device = 1;
                                            }
                                            else
                                            {
                                                Now_Device++;
                                                if (Now_Device == Count_Referee)
                                                {
                                                    Now_Device = Now_Device;
                                                }
                                                else
                                                {
                                                    Now_Device = (Now_Device % Count_Referee);
                                                }
                                            }

                                            if (Flag_Correctness[Now_Device] == "0" || Flag_Performance01[Now_Device] == "0" || Flag_Performance02[Now_Device] == "0" || Flag_Performance03[Now_Device] == "0")
                                            {
                                                break;
                                            }
                                            if (Milestone_Now_Device == Now_Device)//代表全部找完
                                            {
                                                
                                                break;
                                            }
                                        }
                                        #endregion
                                        //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    }
                                }
                                if (Receive_OPcode == "06")
                                {
                                    Flag_Pooling[Receive_DeviceID] = 0;//有收到Ack,就把它歸0繼續往下一台機器
                                    Flag_Performance[Receive_DeviceID] = "1";
                                    Score_Performance[Receive_DeviceID] = Chr(Convert.ToInt32(Data2_Byte1, 16)) + "." + Chr(Convert.ToInt32(Data2_Byte2, 16));
                                    Flag_Correctness[Receive_DeviceID] = "1";
                                    Score_Correctness[Receive_DeviceID] = Chr(Convert.ToInt32(Data1_Byte1, 16)).ToString() + "." + Chr(Convert.ToInt32(Data1_Byte2, 16)).ToString();

                                    Milestone_Now_Device = Now_Device;//設里程碑，若找過一輪都有收到數值了則結束

                                    if (Now_Wait_Device == Receive_DeviceID)
                                    {
                                        //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                        #region 決定下一個Device要去polling誰
                                        if (Count_Referee == 1)//如果只有1台Device
                                        {
                                            Now_Device = 1;//則接下來Polling的Deviece都是1
                                        }
                                        else //若有1個以上的Device
                                        {
                                            Now_Device++;//往下一個Device前進
                                            if (Now_Device == Count_Referee)//由於NowDevice若和機台個數相等mod會變成0
                                            {
                                                Now_Device = Now_Device;//所以若相等則讓他變成加完以後的機號而不做mod
                                            }
                                            else
                                            {
                                                Now_Device = (Now_Device % Count_Referee);//若加完不為機台個數則去mod
                                            }
                                        }

                                        while (Flag_Correctness[Now_Device] == "1" && Flag_Performance01[Now_Device] == "1" && Flag_Performance02[Now_Device] == "1" && Flag_Performance03[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
                                        {
                                            if (Count_Referee == 1)
                                            {
                                                Now_Device = 1;
                                            }
                                            else
                                            {
                                                Now_Device++;
                                                if (Now_Device == Count_Referee)
                                                {
                                                    Now_Device = Now_Device;
                                                }
                                                else
                                                {
                                                    Now_Device = (Now_Device % Count_Referee);
                                                }
                                            }

                                            if (Flag_Correctness[Now_Device] == "0" || Flag_Performance01[Now_Device] == "0" || Flag_Performance02[Now_Device] == "0" || Flag_Performance03[Now_Device] == "0")
                                            {
                                                break;
                                            }
                                            if (Milestone_Now_Device == Now_Device)//代表全部找完
                                            {
                                               
                                                break;
                                            }
                                        }
                                        #endregion
                                        //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    }
                                }
 
                            }
                            #endregion
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }*/

        private void button1_Click(object sender, EventArgs e)
        {
                    
            //serialPort1.Write(CommandConstruct(7, "01"), 0, 16);//serialPort1.Write(Byte16_PCAskForData, 0, 16);
            this.textBox1.AppendText(Environment.NewLine);
            this.textBox1.AppendText("TX : " + ByteToHex(CommandConstruct(7, "01")));
            
            //tb_score1.Text = ByteToHex(CommandConstruct(5, "02"));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void bt_scan_Click(object sender, EventArgs e)
        {
            /*for (int i = 1; i <= 5; i++)//連續發送5次開始以防沒有收到
            {
                Thread.Sleep(50);
                serialPort1.Write(CommandConstruct(0, "04"), 0, 16);
            }*/
            //serialPort1.Write(CommandConstruct(0, "04"), 0, 16); //serialPort1.Write(Byte16_StartRating, 0, 16);
        }

        private void bt_establish_Click(object sender, EventArgs e)
        {
            /*for (int i = 1; i <= 5; i++)//連續發送5次結束以防沒有收到
            {
                Thread.Sleep(50);
                serialPort1.Write(CommandConstruct(0, "05"), 0, 16);
            }*/
            //serialPort1.Write(CommandConstruct(0, "05"), 0, 16);//serialPort1.Write(Byte16_EndRating, 0, 16);
        }

        private void ImportSetting_Click(object sender, EventArgs e)
        {
            //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
            ImportSetting ImportSetting = new ImportSetting();
            ImportSetting.FormClosed += new FormClosedEventHandler(ImportSetting_FormClosed);
            ImportSetting.Show();
            //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
        }
        public static string ParentPath;
        private void ImportSetting_FormClosed(object sender, FormClosedEventArgs e)
        {
            btn_ImportSetting.Text = "載入完畢";
            ImportData = comt.ImportSetting.ImportData;//把exl資料表格載入
            Select_Pose = comt.ImportSetting.Select_Pose;//在ImportSetting時選的要比幾個品勢
            Raw_CompetitorNumber = comt.ImportSetting.Raw_CompetitorNumber;//共有多少位參賽者
            Column_Pose1 = comt.ImportSetting.Column_Pose1;//若選單品勢會有幾行
            Column_Pose2 = comt.ImportSetting.Column_Pose2;//若選雙品勢會有幾行
            Count_Referee = comt.ImportSetting.Count_Referee;//有幾位評審
            ImportData_Score = comt.ImportSetting.ImportData_Score;//把exl分數表格載入    
            //--------------------------------------------------------------------------------
            ExcelFileLocation = comt.ImportSetting.ExcelFileLocation;
            //--------------------------------------------------------------------------------
            ParentPath = comt.ImportSetting.ParentPath;
            #region Set主控端評審
            Color Connect_Yes = Color.Green;//Color.Gray;
            TWS = new TaekwondoSerial.TaekwondoSerial(Count_Referee);
            if (TWS.serialOpen(PortName))
            {
                if (Count_Referee == 1)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                }
                if (Count_Referee == 2)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                    lb_device02.Text = "正\n常";
                    lb_device02.ForeColor = Connect_Yes;
                }
                if (Count_Referee == 3)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                    lb_device02.Text = "正\n常";
                    lb_device02.ForeColor = Connect_Yes;
                    lb_device03.Text = "正\n常";
                    lb_device03.ForeColor = Connect_Yes;
                }
                if (Count_Referee == 4)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                    lb_device02.Text = "正\n常";
                    lb_device02.ForeColor = Connect_Yes;
                    lb_device03.Text = "正\n常";
                    lb_device03.ForeColor = Connect_Yes;
                    lb_device04.Text = "正\n常";
                    lb_device04.ForeColor = Connect_Yes;
                }
                if (Count_Referee == 5)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                    lb_device02.Text = "正\n常";
                    lb_device02.ForeColor = Connect_Yes;
                    lb_device03.Text = "正\n常";
                    lb_device03.ForeColor = Connect_Yes;
                    lb_device04.Text = "正\n常";
                    lb_device04.ForeColor = Connect_Yes;
                    lb_device05.Text = "正\n常";
                    lb_device05.ForeColor = Connect_Yes;
                }
                if (Count_Referee == 6)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                    lb_device02.Text = "正\n常";
                    lb_device02.ForeColor = Connect_Yes;
                    lb_device03.Text = "正\n常";
                    lb_device03.ForeColor = Connect_Yes;
                    lb_device04.Text = "正\n常";
                    lb_device04.ForeColor = Connect_Yes;
                    lb_device05.Text = "正\n常";
                    lb_device05.ForeColor = Connect_Yes;
                    lb_device06.Text = "正\n常";
                    lb_device06.ForeColor = Connect_Yes;
                }
                if (Count_Referee == 7)
                {
                    lb_device01.Text = "正\n常";
                    lb_device01.ForeColor = Connect_Yes;
                    lb_device02.Text = "正\n常";
                    lb_device02.ForeColor = Connect_Yes;
                    lb_device03.Text = "正\n常";
                    lb_device03.ForeColor = Connect_Yes;
                    lb_device04.Text = "正\n常";
                    lb_device04.ForeColor = Connect_Yes;
                    lb_device05.Text = "正\n常";
                    lb_device05.ForeColor = Connect_Yes;
                    lb_device06.Text = "正\n常";
                    lb_device06.ForeColor = Connect_Yes;
                    lb_device07.Text = "正\n常";
                    lb_device07.ForeColor = Connect_Yes;
                }
            }
            //Set初始化Score_Array=========================================================================
            for (int i = 0; i < Constant.ScoringItemNum; i++)
            {
                ScoreArray[i] = new float[Count_Referee + 1];
                for (int j = 0; j < Count_Referee + 1; j++)  // Refree + 1
                {
                    ScoreArray[i][j] = -1;
                }
            }
           
            #endregion
            //Set主控端評審================================================================================
            btn_Demonstrate.Enabled = true;
            btn_Demonstrate.ForeColor = Color.Red;
            //btn_Rating01.Enabled = true;
            if (Select_Pose == "2")//如果選擇比兩個品勢則剛開始的時候，將第2品勢的按鈕先enable = false (就是讓使用者按不到)
            {
                btn_Demonstrate02.Enabled = false;
                btn_Demonstrate02_Start.Enabled = false;
                btn_Rating02.Enabled = false;
                //---------------
                btn_Demonstrate01_Start.Enabled = false;
                btn_Rating01.Enabled = false;
            }
            else//如果只選擇比單一品勢則隱藏第二品勢按鈕
            {
                //隱藏第二品勢資訊=================================================
                btn_Demonstrate02.Visible = false;
                btn_Rating02.Visible = false;
                btn_Demonstrate02_Start.Visible = false;
                pnl_background02.Visible = false;
                pnl_background04.Visible = false;
                lb_Pose02_ScoreTotal_Title.Visible = false;
                pnl_Pose02_ScoreTotal.Visible = false;
                lb_Pose02_ScoreTotal.Visible = false;                
                //-------------------------------------------
                lb_Pose02_ScoreTotal_Title.Visible = false;
                pnl_Pose02_ScoreTotal.Visible = false;
                lb_Pose02_ScoreTotal.Visible = false;

                lb_ScoreAverage_Title.Visible = false;
                pnl_ScoreAverage.Visible = false;
                lb_ScoreAverage.Visible = false;
                //隱藏第二品勢資訊=================================================
            }

            btn_Next.Enabled = true;
            btn_ReNew.Enabled = true;
            btn_Abstain.Enabled = true;
            //--------------------------------------------------------------------------------
            if (ImportData != null)
            {
                lb_NowPlayer.Text = ImportData[15, Now_Player];
                lb_NowPlayer_Name.Text = ImportData[2, Now_Player];
                lb_NowPlayer_Group.Text = ImportData[0, Now_Player];
                lb_NowPlayer_Unit.Text = ImportData[1, Now_Player];
                lb_NowPlayer_Match.Text = ImportData[3, Now_Player];
            }
            

            //----------------------------------------------------------------------------------------
            timer_CheckRatingState.Enabled = true;//開始對參賽者評分狀態做掃描顯示
            IsOpen_ImportSetting = true;//代表已經載入過資料
            btn_ImportSetting.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            WriteExcelAll();
        }

        public static void WriteExcelAll()
        {
            #region 將全部資料寫入Excel，選手資料寫入sheet1，分數寫入Sheet2
            int i;
            string PathFile = ExcelFileLocation;// Directory.GetCurrentDirectory() + @"\Data.xlsx";
            _Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook myBook = myExcel.Workbooks.Open(PathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet mySheet = (Worksheet)myBook.Worksheets[1];
            //string ooo = ImportData[6, 1];
            if (Select_Pose == "1")
            {
                for (i = 1; i <= (Raw_CompetitorNumber + 1); i++)
                {
                    #region 新版
                    myExcel.Cells[i, 1] = ImportData[17, i - 1];//組別代號
                    myExcel.Cells[i, 2] = ImportData[0, i - 1];//組別名稱
                    myExcel.Cells[i, 3] = ImportData[1, i - 1];//單位
                    myExcel.Cells[i, 4] = ImportData[2, i - 1];//姓名
                    myExcel.Cells[i, 5] = ImportData[15, i - 1];//籤號
                    myExcel.Cells[i, 6] = ImportData[3, i - 1];//初賽或複賽
                    myExcel.Cells[i, 7] = ImportData[4, i - 1];//第一品勢名稱
                    myExcel.Cells[i, 8] = ImportData[5, i - 1];//第二品勢名稱
                    myExcel.Cells[i, 9] = ImportData[18, i - 1];//國旗
                    myExcel.Cells[i, 10] = ImportData[13, i - 1];//排名
                    myExcel.Cells[i, 11] = ImportData[14, i - 1];//名次
                    myExcel.Cells[i, 12] = ImportData[6, i - 1];//第一品勢正確性分數
                    myExcel.Cells[i, 13] = ImportData[7, i - 1];//第一品勢表現性分數
                    myExcel.Cells[i, 14] = ImportData[8, i - 1];//第一品勢總分
                    myExcel.Cells[i, 15] = ImportData[19, i - 1];//第一品勢原始總分
                    myExcel.Cells[i, 16] = ImportData[9, i - 1];//第二品勢正確性分數
                    myExcel.Cells[i, 17] = ImportData[10, i - 1];//第二品勢表現性分數
                    //-----------------------------------------------------------------
                    myExcel.Cells[i, 18] = ImportData[11, i - 1];//第二品勢總分
                    myExcel.Cells[i, 19] = ImportData[20, i - 1];//第二品勢原始總分
                    //-----------------------------------------------------------------
                    myExcel.Cells[i, 20] = ImportData[12, i - 1];//總平均
                    myExcel.Cells[i, 21] = ImportData[21, i - 1];//原始總分
                    myExcel.Cells[i, 22] = ImportData[16, i - 1];//棄權
                    #endregion
                }
                /*
                mySheet = (Worksheet)myBook.Worksheets[2];
                mySheet.Activate();//設工作表焦點

                //ImportData_Score = new string[29, (Height)];
                for (i = 1; i <= (Raw_CompetitorNumber + 1); i++)
                {
                    myExcel.Cells[i, 1] = ImportData_Score[0, i - 1];
                    myExcel.Cells[i, 2] = ImportData_Score[1, i - 1];
                    myExcel.Cells[i, 3] = ImportData_Score[2, i - 1];
                    myExcel.Cells[i, 4] = ImportData_Score[3, i - 1];
                    myExcel.Cells[i, 5] = ImportData_Score[4, i - 1];
                    myExcel.Cells[i, 6] = ImportData_Score[5, i - 1];
                    myExcel.Cells[i, 7] = ImportData_Score[6, i - 1];
                    myExcel.Cells[i, 8] = ImportData_Score[7, i - 1];
                    myExcel.Cells[i, 9] = ImportData_Score[8, i - 1];
                    myExcel.Cells[i, 10] = ImportData_Score[9, i - 1];
                    myExcel.Cells[i, 11] = ImportData_Score[10, i - 1];
                    myExcel.Cells[i, 12] = ImportData_Score[11, i - 1];
                    myExcel.Cells[i, 13] = ImportData_Score[12, i - 1];
                    myExcel.Cells[i, 14] = ImportData_Score[13, i - 1];
                    myExcel.Cells[i, 15] = ImportData_Score[14, i - 1];
                    myExcel.Cells[i, 16] = ImportData_Score[15, i - 1];
                    myExcel.Cells[i, 17] = ImportData_Score[16, i - 1];
                    myExcel.Cells[i, 18] = ImportData_Score[17, i - 1];
                    myExcel.Cells[i, 19] = ImportData_Score[18, i - 1];
                    myExcel.Cells[i, 20] = ImportData_Score[19, i - 1];
                    myExcel.Cells[i, 21] = ImportData_Score[20, i - 1];
                    myExcel.Cells[i, 22] = ImportData_Score[21, i - 1];
                    myExcel.Cells[i, 23] = ImportData_Score[22, i - 1];
                    myExcel.Cells[i, 24] = ImportData_Score[23, i - 1];
                    myExcel.Cells[i, 25] = ImportData_Score[24, i - 1];
                    myExcel.Cells[i, 26] = ImportData_Score[25, i - 1];
                    myExcel.Cells[i, 27] = ImportData_Score[26, i - 1];
                    myExcel.Cells[i, 28] = ImportData_Score[27, i - 1];
                    myExcel.Cells[i, 29] = ImportData_Score[28, i - 1];
                }
                */
            }
            else if (Select_Pose == "2")
            {
                for (i = 1; i <= (Raw_CompetitorNumber + 1); i++)
                {
                    #region 原本
                    /*
                    myExcel.Cells[i, 1] = ImportData[0, i - 1];//組別
                    myExcel.Cells[i, 2] = ImportData[1, i - 1];//單位
                    myExcel.Cells[i, 3] = ImportData[2, i - 1];//姓名
                    myExcel.Cells[i, 4] = ImportData[3, i - 1];//初賽或複賽
                    myExcel.Cells[i, 5] = ImportData[4, i - 1];//第一品勢名稱
                    myExcel.Cells[i, 6] = ImportData[5, i - 1];//第二品勢名稱
                    myExcel.Cells[i, 7] = ImportData[6, i - 1];//第一品勢正確性分數
                    myExcel.Cells[i, 8] = ImportData[7, i - 1];//第一品勢表現性分數
                    myExcel.Cells[i, 9] = ImportData[8, i - 1];//第一品勢平均
                    myExcel.Cells[i, 10] = ImportData[9, i - 1];//第二品勢正確性分數
                    myExcel.Cells[i, 11] = ImportData[10, i - 1];//第二品勢表現性分數
                    myExcel.Cells[i, 12] = ImportData[11, i - 1];//第二品勢平均
                    myExcel.Cells[i, 13] = ImportData[12, i - 1];//總平均
                    myExcel.Cells[i, 14] = ImportData[13, i - 1];//排名(數字)
                    myExcel.Cells[i, 15] = ImportData[14, i - 1];//名次(國字)
                    myExcel.Cells[i, 16] = ImportData[15, i - 1];//編號
                    myExcel.Cells[i, 17] = ImportData[16, i - 1];//棄權
                    */
                    /*
                    myExcel.Cells[i, 1] = ImportData[17, i - 1];//組別代號
                    myExcel.Cells[i, 2] = ImportData[0, i - 1];//組別
                    myExcel.Cells[i, 3] = ImportData[1, i - 1];//單位
                    myExcel.Cells[i, 4] = ImportData[2, i - 1];//姓名
                    myExcel.Cells[i, 5] = ImportData[15, i - 1];//編號
                    myExcel.Cells[i, 6] = ImportData[3, i - 1];//初賽或複賽
                    myExcel.Cells[i, 7] = ImportData[4, i - 1];//第一品勢名稱
                    myExcel.Cells[i, 8] = ImportData[5, i - 1];//第二品勢名稱
                    myExcel.Cells[i, 9] = ImportData[18, i - 1];//國旗
                    myExcel.Cells[i, 10] = ImportData[13, i - 1];//排名(數字)
                    myExcel.Cells[i, 11] = ImportData[14, i - 1];//名次(國字)
                    myExcel.Cells[i, 12] = ImportData[6, i - 1];//第一品勢正確性分數
                    myExcel.Cells[i, 13] = ImportData[7, i - 1];//第一品勢表現性分數
                    myExcel.Cells[i, 14] = ImportData[8, i - 1];//第一品勢平均
                    myExcel.Cells[i, 15] = ImportData[9, i - 1];//第二品勢正確性分數
                    myExcel.Cells[i, 16] = ImportData[10, i - 1];//第二品勢表現性分數
                    myExcel.Cells[i, 17] = ImportData[11, i - 1];//第二品勢平均
                    //-----------------------------------------------------------------
                    myExcel.Cells[i, 18] = ImportData[12, i - 1];//總平均
                    myExcel.Cells[i, 19] = ImportData[16, i - 1];//棄權 
                    */

                    #endregion

                    #region 新版
                    myExcel.Cells[i, 1] = ImportData[17, i - 1];//組別代號
                    myExcel.Cells[i, 2] = ImportData[0, i - 1];//組別名稱
                    myExcel.Cells[i, 3] = ImportData[1, i - 1];//單位
                    myExcel.Cells[i, 4] = ImportData[2, i - 1];//姓名
                    myExcel.Cells[i, 5] = ImportData[15, i - 1];//籤號
                    myExcel.Cells[i, 6] = ImportData[3, i - 1];//初賽或複賽
                    myExcel.Cells[i, 7] = ImportData[4, i - 1];//第一品勢名稱
                    myExcel.Cells[i, 8] = ImportData[5, i - 1];//第二品勢名稱
                    myExcel.Cells[i, 9] = ImportData[18, i - 1];//國旗
                    myExcel.Cells[i, 10] = ImportData[13, i - 1];//排名
                    myExcel.Cells[i, 11] = ImportData[14, i - 1];//名次
                    myExcel.Cells[i, 12] = ImportData[6, i - 1];//第一品勢正確性分數
                    myExcel.Cells[i, 13] = ImportData[7, i - 1];//第一品勢表現性分數
                    myExcel.Cells[i, 14] = ImportData[8, i - 1];//第一品勢總分
                    myExcel.Cells[i, 15] = ImportData[19, i - 1];//第一品勢原始總分
                    myExcel.Cells[i, 16] = ImportData[9, i - 1];//第二品勢正確性分數
                    myExcel.Cells[i, 17] = ImportData[10, i - 1];//第二品勢表現性分數
                    //-----------------------------------------------------------------
                    myExcel.Cells[i, 18] = ImportData[11, i - 1];//第二品勢總分
                    myExcel.Cells[i, 19] = ImportData[20, i - 1];//第二品勢原始總分
                    //-----------------------------------------------------------------
                    myExcel.Cells[i, 20] = ImportData[12, i - 1];//總平均
                    myExcel.Cells[i, 21] = ImportData[21, i - 1];//原始總分
                    myExcel.Cells[i, 22] = ImportData[16, i - 1];//棄權
                    #endregion
                }
                /*
                mySheet = (Worksheet)myBook.Worksheets[2];
                mySheet.Activate();//設工作表焦點

                for (i = 1; i <= (Raw_CompetitorNumber + 1); i++)
                {
                    myExcel.Cells[i, 1] = ImportData_Score[0, i - 1];
                    myExcel.Cells[i, 2] = ImportData_Score[1, i - 1];
                    myExcel.Cells[i, 3] = ImportData_Score[2, i - 1];
                    myExcel.Cells[i, 4] = ImportData_Score[3, i - 1];
                    myExcel.Cells[i, 5] = ImportData_Score[4, i - 1];
                    myExcel.Cells[i, 6] = ImportData_Score[5, i - 1];
                    myExcel.Cells[i, 7] = ImportData_Score[6, i - 1];
                    myExcel.Cells[i, 8] = ImportData_Score[7, i - 1];
                    myExcel.Cells[i, 9] = ImportData_Score[8, i - 1];
                    myExcel.Cells[i, 10] = ImportData_Score[9, i - 1];
                    myExcel.Cells[i, 11] = ImportData_Score[10, i - 1];
                    myExcel.Cells[i, 12] = ImportData_Score[11, i - 1];
                    myExcel.Cells[i, 13] = ImportData_Score[12, i - 1];
                    myExcel.Cells[i, 14] = ImportData_Score[13, i - 1];
                    myExcel.Cells[i, 15] = ImportData_Score[14, i - 1];
                    myExcel.Cells[i, 16] = ImportData_Score[15, i - 1];
                    myExcel.Cells[i, 17] = ImportData_Score[16, i - 1];
                    myExcel.Cells[i, 18] = ImportData_Score[17, i - 1];
                    myExcel.Cells[i, 19] = ImportData_Score[18, i - 1];
                    myExcel.Cells[i, 20] = ImportData_Score[19, i - 1];
                    myExcel.Cells[i, 21] = ImportData_Score[20, i - 1];
                    myExcel.Cells[i, 22] = ImportData_Score[21, i - 1];
                    myExcel.Cells[i, 23] = ImportData_Score[22, i - 1];
                    myExcel.Cells[i, 24] = ImportData_Score[23, i - 1];
                    myExcel.Cells[i, 25] = ImportData_Score[24, i - 1];
                    myExcel.Cells[i, 26] = ImportData_Score[25, i - 1];
                    myExcel.Cells[i, 27] = ImportData_Score[26, i - 1];
                    myExcel.Cells[i, 28] = ImportData_Score[27, i - 1];
                    myExcel.Cells[i, 29] = ImportData_Score[28, i - 1];
                }
                */
            }

            //button2.Text = ((Range)mySheet.Cells[1, 1]).Text.ToString();
            //button2.Text = "Width:" + Width.ToString() + "Height:" + Height.ToString() + "有" + (Height - 1).ToString() + "個選手";
            /*
            wbook.SaveAs(PathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                            , XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            */
            myBook.Save();
            myBook.Close(false, Type.Missing, Type.Missing);//關閉Excel
            myExcel.Quit();//釋放Excel資源            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            myExcel = null;
            mySheet = null;
            //Range = null;
            myBook = null;
            GC.Collect();

            #endregion
        }

        private void btn_form2_Click(object sender, EventArgs e)
        {
            /*
            Form2 Form2 = new Form2();
            Form2.Show();
            */
            Rank Rank = new Rank();
            Rank.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Keyboard Keyboard = new Keyboard();
            Keyboard.Show();
        }

        private void btn_DeviceNameCheck_Click(object sender, EventArgs e)
        {
            /*
            Result Result = new Result();
            Result.Show();
            */
            /*
            string CRC_StartRating;
            string CRC_EndRating;
            string CRC_PCAskForData;
            string CRC_RoutineAck;

            byte[] Byte12_StartRating;
            byte[] Byte12_EndRating;
            byte[] Byte12_PCAskForData;
            byte[] Byte12_RoutineAck;

            byte[] Byte16_StartRating;
            byte[] Byte16_EndRating;
            byte[] Byte16_PCAskForData;
            byte[] Byte16_RoutineAck;
            */

            DeviceName = tb_DeviceName.Text;
            //StartRating///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region StartRating
            Byte12_StartRating = new byte[12] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x34", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16)
            };
            CRC_StartRating = String.Format("{0,4:X}", CRC_any(Byte12_StartRating, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
            CRC_StartRating = CRC_StartRating.Replace(" ", "0");
            Byte16_StartRating = new byte[16] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x34", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                //----------------------------
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x"+CRC_StartRating.Substring(0,2), 16),
                Convert.ToByte("0x"+CRC_StartRating.Substring(2,2), 16)
            };
            #endregion
            //StartRating///////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //EndRating/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region EndRating
            Byte12_EndRating = new byte[12] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x35", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16)
            };
            CRC_EndRating = String.Format("{0,4:X}", CRC_any(Byte12_EndRating, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
            CRC_EndRating = CRC_EndRating.Replace(" ", "0");
            Byte16_EndRating = new byte[16] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x35", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                //----------------------------
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x"+CRC_EndRating.Substring(0,2), 16),
                Convert.ToByte("0x"+CRC_EndRating.Substring(2,2), 16)
            };
            #endregion
            //EndRating/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //PCAskForData//////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region PCAskForData
            Byte12_PCAskForData = new byte[12] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x31", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16)
            };
            CRC_PCAskForData = String.Format("{0,4:X}", CRC_any(Byte12_PCAskForData, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
            CRC_PCAskForData = CRC_PCAskForData.Replace(" ", "0");
            Byte16_PCAskForData = new byte[16] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x31", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                //----------------------------
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x"+CRC_PCAskForData.Substring(0,2), 16),
                Convert.ToByte("0x"+CRC_PCAskForData.Substring(2,2), 16)
            };
            #endregion
            //PCAskForData//////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //RoutineAck////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region RoutineAck
            Byte12_RoutineAck = new byte[12] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x32", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16)
            };
            CRC_RoutineAck = String.Format("{0,4:X}", CRC_any(Byte12_RoutineAck, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
            CRC_RoutineAck = CRC_RoutineAck.Replace(" ", "0");
            Byte16_RoutineAck = new byte[16] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x30", 16),//Opcode
                Convert.ToByte("0x32", 16),//Opcode
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                Convert.ToByte("0x30", 16),
                //----------------------------
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x00", 16),
                Convert.ToByte("0x"+CRC_RoutineAck.Substring(0,2), 16),
                Convert.ToByte("0x"+CRC_RoutineAck.Substring(2,2), 16)
            };
            #endregion
            //RoutineAck////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            btn_DeviceNameCheck.Visible = false;
            tb_DeviceName.Enabled = false;
        }
        
        private void btn_Demonstrate_Click(object sender, EventArgs e)
        {
            if (lb_Pose01_ScoreTotal.Text == "")
            {
                btn_Demonstrate.ForeColor = Color.White;
                btn_Demonstrate.Enabled = false;
                if (IsOpen_Rank == true)
                {
                    //進入Demonstrate畫面前關閉Rank畫面/////////////////////////////////               
                    comt.Result.Rank.Close();
                    //進入Demonstrate畫面前關閉Rank畫面/////////////////////////////////               
                }
                if (IsOpen_Rank_ByRating == true)
                {
                    comt.Rating.Rank.Close();
                }
                //IsOpen_Rank = false;//因為已經開啟Demonstrate的視窗了，所以將Rank關掉後將IsOpen_Rank設為false
                //當按下表演的時候，必須要能夠 讓 計時開始的按鍵 和 評分的按鍵 能按
                btn_Demonstrate01_Start.Enabled = true;
                btn_Demonstrate01_Start.ForeColor = Color.Red;
                btn_Rating01.Enabled = false;
                //當按下表演的時候，必須要能夠 讓 計時開始的按鍵 和 評分的按鍵 能按

                //讓評分的地方知道表演視窗已開啟，以防日後評分視窗去關掉一個未開啟的視窗//////////////////////////////////////////////////
                IsOpen_Demonstrate = true;
                //讓評分的地方知道表演視窗已開啟，以防日後評分視窗去關掉一個未開啟的視窗//////////////////////////////////////////////////

                Now_Pose = 1;

                //廣播各Device說，已經開始表演了，可以打開機器讓評審一邊看比賽一邊評分了(廣播5次，Device有收到就有收到沒收到就算了)/////////////
                /*for (int i = 1; i <= 5; i++)//連續發送5次開始以防沒有收到
                {
                    Thread.Sleep(50);
                    serialPort1.Write(CommandConstruct(0, "04"), 0, 16);
                }*/
                #region mark20130510_1604
                /*
                for (int i = 1; i <= 5; i++)
                {
                    Thread.Sleep(50);
                    for (int j = 1; j <= Count_Referee; j++)
                    serialPort1.Write(CommandConstruct(i, "04"), 0, 16);
                }
            */
                #endregion
                //廣播各Device說，已經開始表演了，可以打開機器讓評審一邊看比賽一邊評分了(廣播5次，Device有收到就有收到沒收到就算了///////////// 

                //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                Demonstrate = new Demonstrate();
                Demonstrate.FormClosed += new FormClosedEventHandler(Demonstrate_FormClosed);
                Demonstrate.Show();
                //偵測子畫面關閉，父畫面做事情--------------------------------------------------------- 
                if (btn_OPEN_COM.Text == "斷線" && Count_Referee > 0)
                {
                    for (int i = 0; i < 5; i++)
                        //TWS.clearAllReferee();
                        TWS.clearAllReferee();
                }
                TWS.clearAllReferee();
            }
            else
            {
                MessageBox.Show("此人已平分過了，若要重新評分，請選擇重新評分", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_Demonstrate02_Click(object sender, EventArgs e)
        {
            //EasonFlow///////////////////////////////////////////////////
            //把自己變白
            btn_Demonstrate02.ForeColor = Color.White;
            btn_Demonstrate02.Enabled = false;
            //把下一個變紅
            btn_Demonstrate02_Start.ForeColor = Color.Red;
            btn_Demonstrate02_Start.Enabled = true;
            //EasonFlow///////////////////////////////////////////////////

            if (ImportData[8, Now_Player] == "")//第一品勢尚未評分完畢，則不允許進行第二品勢評分
            {
                MessageBox.Show("請先將第一品勢，進行評分", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else 
            {
                if (IsOpen_Rank == true)
                {
                    //進入Demonstrate畫面前關閉Rank畫面/////////////////////////////////               
                    comt.Result.Rank.Close();
                    //進入Demonstrate畫面前關閉Rank畫面/////////////////////////////////               
                }
                //IsOpen_Rank = false;//因為已經開啟Demonstrate的視窗了，所以將Rank關掉後將IsOpen_Rank設為false

                if (IsOpen_Rating == true)
                {
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////               
                    Rating.Close();
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                }
                //當按下表演的時候，必須要能夠 讓 計時開始的按鍵 和 評分的按鍵 能按
                btn_Demonstrate02_Start.Enabled = true;
                //btn_Rating02.Enabled = true;
                //當按下表演的時候，必須要能夠 讓 計時開始的按鍵 和 評分的按鍵 能按


                if (Select_Pose == "1")
                {
                    MessageBox.Show("此次比賽僅為單一品勢比賽", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    //讓評分的地方知道表演視窗已開啟，以防日後評分視窗去關掉一個未開啟的視窗//////////////////////////////////////////////////
                    IsOpen_Demonstrate = true;
                    //讓評分的地方知道表演視窗已開啟，以防日後評分視窗去關掉一個未開啟的視窗//////////////////////////////////////////////////

                    Now_Pose = 2;

                    //廣播各Device說，已經開始表演了，可以打開機器讓評審一邊看比賽一邊評分了(廣播5次，Device有收到就有收到沒收到就算了)/////////////
                    /*for (int i = 1; i <= 5; i++)
                    {
                        Thread.Sleep(50);
                        serialPort1.Write(CommandConstruct(0, "04"), 0, 16);
                    }*/

                    //廣播各Device說，已經開始表演了，可以打開機器讓評審一邊看比賽一邊評分了(廣播5次，Device有收到就有收到沒收到就算了///////////// 

                    //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                    Demonstrate02 = new Demonstrate();
                    Demonstrate02.FormClosed += new FormClosedEventHandler(Demonstrate02_FormClosed);
                    Demonstrate02.Show();
                    //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                }
                if (btn_OPEN_COM.Text == "斷線" && Count_Referee > 0)
                {
                    for (int i = 0; i < 5; i++)
                        //TWS.clearAllReferee();
                        TWS.clearAllReferee();
                }
                TWS.clearAllReferee();
            }            
            
        }

        Demonstrate Demonstrate02;
        private void Demonstrate_FormClosed(object sender, FormClosedEventArgs e)
        {//關閉螢幕
            IsOpen_Demonstrate = false;
        }

        private void Demonstrate02_FormClosed(object sender, FormClosedEventArgs e)
        {//關閉螢幕
            IsOpen_Demonstrate = false;
        }
        //開始評分
        private void btn_Rating01_Click(object sender, EventArgs e)
        {
            //EasonFlow///////////////////////////////////////////////////
            //把自己變白
            
            btn_Rating01.ForeColor = Color.White;
            btn_Rating01.Enabled = false;
            /*
            //把下一個變紅
            btn_Demonstrate02.ForeColor = Color.Red;
            btn_Demonstrate02.Enabled = true;*/
            //EasonFlow///////////////////////////////////////////////////
            IsOpen_Rating = true;//評分視窗開啟
            //btn_Demonstrate02.Enabled = true;
            if (IsOpen_Demonstrate == true)
            {
                //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                RemoteControl_TimerStart = false;
                comt.Demonstrate.timer.Enabled = false;                
                Demonstrate.Close();
                //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
            }            
            //收值陣列初始化/////////////////////////////////////////////////////////////////////
            Flag_Correctness = new string[8];
            Flag_Performance01 = new string[8];
            Flag_Performance02 = new string[8];
            Flag_Performance03 = new string[8];

            Score_Correctness = new string[8];
            Score_Performance01 = new string[8];
            Score_Performance02 = new string[8];
            Score_Performance03 = new string[8];
            Flag_Pooling = new int[8];

            for (int i = 1; i <= Count_Referee; i++) 
            {
                Flag_Correctness[i] = "0";
                Flag_Performance01[i] = "0";
                Flag_Performance02[i] = "0";
                Flag_Performance03[i] = "0";
                Score_Correctness[i] = "0";
                Score_Performance01[i] = "0";
                Score_Performance02[i] = "0";
                Score_Performance03[i] = "0";
                Flag_Pooling[i] = 0;
            }
            Now_Pose = 1;//告知現在評分的是第1品勢
            Now_Device = 1;//都從第一個Device開始收資料
            //收值陣列初始化/////////////////////////////////////////////////////////////////////
            
            
            //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
            Rating = new Rating();
            Rating.FormClosed += new FormClosedEventHandler(Rating_FormClosed);
            Rating.Show();
            //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
            timer_Scoring.Start();
          
        }

        private void btn_Rating02_Click(object sender, EventArgs e)
        {
            //EasonFlow///////////////////////////////////////////////////
            //把自己變白
            btn_Rating02.ForeColor = Color.Yellow;
            btn_Rating02.Enabled = false;
            //EasonFlow//////////////////////////////////////////////////
            IsOpen_Rating = true;//評分視窗開啟
            if (Select_Pose == "1")
            {                
                MessageBox.Show("此次比賽僅為單一品勢比賽", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Now_Pose = 2;
                if (IsOpen_Demonstrate == true)
                {
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                    comt.Demonstrate.timer.Enabled = false;
                    RemoteControl_TimerStart = false;
                    Demonstrate02.Close();
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                }
                //收值陣列初始化/////////////////////////////////////////////////////////////////////
                Flag_Correctness = new string[8];
                Flag_Performance01 = new string[8];
                Flag_Performance02 = new string[8];
                Flag_Performance03 = new string[8];
                Score_Correctness = new string[8];
                Score_Performance01 = new string[8];
                Score_Performance02 = new string[8];
                Score_Performance03 = new string[8];
                Flag_Pooling = new int[8];

                for (int i = 1; i <= Count_Referee; i++)
                {
                    Flag_Correctness[i] = "0";
                    Flag_Performance01[i] = "0";
                    Flag_Performance02[i] = "0";
                    Flag_Performance03[i] = "0";
                    Score_Correctness[i] = "0";
                    Score_Performance01[i] = "0";
                    Score_Performance02[i] = "0";
                    Score_Performance03[i] = "0";
                    Flag_Pooling[i] = 0;
                }
                Now_Pose = 2;//告知現在評分的是第2品勢
                Now_Device = 1;//都從第一個Device開始收資料
                //收值陣列初始化/////////////////////////////////////////////////////////////////////

                //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                Rating = new Rating();
                Rating.FormClosed += new FormClosedEventHandler(Rating_FormClosed);
                Rating.Show();
                //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                timer_Scoring.Start();
              
            }

        }

        private void Rating_FormClosed(object sender, FormClosedEventArgs e)
        {//關閉螢幕，將資料分別寫入陣列及Excel檔
            IsOpen_Rating = false;//評分視窗關閉
            if (Now_Pose == 1)
            {
                //WriteExcelAll();
            }
            else if (Now_Pose == 2)
            {
                //WriteExcelAll();
            }
            //button2.Text = "關閉了視窗";            
        }

        private void btn_Next_Click(object sender, EventArgs e)
        {
            
            //MessageBox.Show("該選手尚未評分完畢", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (IsWritng_Rank == true)
            {
                MessageBox.Show("檔案寫入中請稍後再試。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else 
            {
                DialogResult MsgBoxResult1;//設置對話框的返回值
                MsgBoxResult1 = MessageBox.Show("確認進行至下一位選手?",//對話框的顯示內容 
                "系統訊息",//對話框的標題 
                MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                if (MsgBoxResult1 == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                {
                    if (Select_Pose == "1")
                    {
                        if (ImportData[8, Now_Player] == "")
                        {
                            MessageBox.Show("該選手尚未評分完畢", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            //EasonFlow/////////////////////////////////////////////////////////
                            btn_Demonstrate.Enabled = true;
                            btn_Demonstrate.ForeColor = Color.Red;
                            //EasonFlow/////////////////////////////////////////////////////////
                            btn_Demonstrate02.Enabled = false;
                            btn_Demonstrate02_Start.Enabled = false;
                            btn_Rating01.Enabled = false;
                            btn_Rating02.Enabled = false;
                            if (Now_Player == Raw_CompetitorNumber)
                            {
                                //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                DialogResult MsgBoxResult;//設置對話框的返回值
                                MsgBoxResult = MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者, " + "確認匯出?",//對話框的顯示內容 //"確認匯出並關閉視窗?"
                                "系統訊息",//對話框的標題 
                                MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                                MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                                MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                                if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                                {
                                    #region 選擇匯出
                                    RankForAllPlayer();
                                    comt.Form1.WriteExcelAll();
                                    WriteExcelExport();//SortExcelFile//匯出
                                    //確認是否上傳?----------------------------------------------------------------------
                                    #region 確認是否上傳?
                                    DialogResult MsgBoxResult_Upload;
                                    MsgBoxResult_Upload = MessageBox.Show("請問是否上傳",
                                    "系統訊息",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button2);
                                    if (MsgBoxResult_Upload == DialogResult.Yes)
                                    {
                                        FTPUpload FTPUpload = new FTPUpload();//選擇上傳開起上傳頁面
                                        FTPUpload.Show();
                                    }
                                    if (MsgBoxResult_Upload == DialogResult.No)
                                    {
                                        #region 選擇不上傳，尋問使用者是否關閉頁面
                                        DialogResult MsgBoxResult_Close;
                                        MsgBoxResult_Close = MessageBox.Show("請問是否關閉程式?",
                                        "系統訊息",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Exclamation,
                                        MessageBoxDefaultButton.Button2);
                                        if (MsgBoxResult_Close == DialogResult.Yes)
                                        {
                                            Environment.Exit(Environment.ExitCode);
                                        }
                                        if (MsgBoxResult_Close == DialogResult.No)
                                        {
                                        }
                                        #endregion
                                    }
                                    #endregion
                                    //確認是否上傳?----------------------------------------------------------------------
                                    //this.Close();
                                    #endregion
                                }
                                if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                                {
                                    #region 選擇不匯出(ie,不會排序)
                                    DialogResult MsgBoxResult_Close;
                                    MsgBoxResult_Close = MessageBox.Show("請問是否關閉程式?",
                                    "系統訊息",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button2);
                                    if (MsgBoxResult_Close == DialogResult.Yes)
                                    {
                                        Environment.Exit(Environment.ExitCode);//不匯出直接關閉程式
                                    }
                                    if (MsgBoxResult_Close == DialogResult.No)
                                    {
                                        //不匯出，也不要關閉程式讓使用者繼續留在操控端畫面
                                    }
                                    //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                                    //label1.Text = " 你選擇了按下”No“的按鈕！";
                                    //MessageBox.Show("那你再考慮考慮");
                                    #endregion
                                }
                            }
                            else
                            {
                                Now_Player++;
                            }
                            lb_NowPlayer.Text = ImportData[15, Now_Player];
                            lb_NowPlayer_Name.Text = ImportData[2, Now_Player];
                            lb_NowPlayer_Group.Text = ImportData[0, Now_Player];
                            lb_NowPlayer_Unit.Text = ImportData[1, Now_Player];
                            lb_NowPlayer_Match.Text = ImportData[3, Now_Player];
                        }
                    }

                    if (Select_Pose == "2")
                    {

                        if (ImportData[8, Now_Player] == "" || ImportData[11, Now_Player] == "")
                        {
                            MessageBox.Show("該選手尚未評分完畢", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            //EasonFlow/////////////////////////////////////////////////////////
                            btn_Demonstrate.Enabled = true;
                            btn_Demonstrate.ForeColor = Color.Red;
                            //EasonFlow/////////////////////////////////////////////////////////
                            btn_Demonstrate02.Enabled = false;
                            btn_Demonstrate02_Start.Enabled = false;
                            btn_Rating01.Enabled = false;
                            btn_Rating02.Enabled = false;
                            if (Now_Player == Raw_CompetitorNumber)
                            {
                                #region mark 原版
                                /*
                                //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                DialogResult MsgBoxResult;//設置對話框的返回值
                                MsgBoxResult = MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者, " + "確認匯出?",//對話框的顯示內容 
                                "系統訊息",//對話框的標題 
                                MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                                MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                                MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                                if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                                {
                                    FTPUpload FTPUpload = new FTPUpload();
                                    FTPUpload.Show();
                                    WriteExcelExport();
                                    //this.Close();
                                }
                                if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                                {
                                    //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                                    //label1.Text = " 你選擇了按下”No“的按鈕！";
                                    //MessageBox.Show("那你再考慮考慮");
                                }
                                */
                                #endregion

                                //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                DialogResult MsgBoxResult;//設置對話框的返回值
                                MsgBoxResult = MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者, " + "確認匯出?",//對話框的顯示內容 //"確認匯出並關閉視窗?"
                                "系統訊息",//對話框的標題 
                                MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                                MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                                MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                                if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                                {
                                    #region 選擇匯出
                                    RankForAllPlayer();
                                    comt.Form1.WriteExcelAll();
                                    WriteExcelExport();//SortExcelFile//匯出
                                    //確認是否上傳?----------------------------------------------------------------------
                                    #region 確認是否上傳?
                                    DialogResult MsgBoxResult_Upload;
                                    MsgBoxResult_Upload = MessageBox.Show("請問是否上傳",
                                    "系統訊息",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button2);
                                    if (MsgBoxResult_Upload == DialogResult.Yes)
                                    {
                                        FTPUpload FTPUpload = new FTPUpload();//選擇上傳開起上傳頁面
                                        FTPUpload.Show();
                                    }
                                    if (MsgBoxResult_Upload == DialogResult.No)
                                    {
                                        #region 選擇不上傳，尋問使用者是否關閉頁面
                                        DialogResult MsgBoxResult_Close;
                                        MsgBoxResult_Close = MessageBox.Show("請問是否關閉程式?",
                                        "系統訊息",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Exclamation,
                                        MessageBoxDefaultButton.Button2);
                                        if (MsgBoxResult_Close == DialogResult.Yes)
                                        {
                                            Environment.Exit(Environment.ExitCode);
                                        }
                                        if (MsgBoxResult_Close == DialogResult.No)
                                        {
                                        }
                                        #endregion
                                    }
                                    #endregion
                                    //確認是否上傳?----------------------------------------------------------------------
                                    //this.Close();
                                    #endregion
                                }
                                if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                                {
                                    #region 選擇不匯出(ie,不會排序)
                                    DialogResult MsgBoxResult_Close;
                                    MsgBoxResult_Close = MessageBox.Show("請問是否關閉程式?",
                                    "系統訊息",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button2);
                                    if (MsgBoxResult_Close == DialogResult.Yes)
                                    {
                                        Environment.Exit(Environment.ExitCode);//不匯出直接關閉程式
                                    }
                                    if (MsgBoxResult_Close == DialogResult.No)
                                    {
                                        //不匯出，也不要關閉程式讓使用者繼續留在操控端畫面
                                    }
                                    //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                                    //label1.Text = " 你選擇了按下”No“的按鈕！";
                                    //MessageBox.Show("那你再考慮考慮");
                                    #endregion
                                }                                
                            }
                            else
                            {
                                Now_Player++;
                            }
                            lb_NowPlayer.Text = ImportData[15, Now_Player];
                            lb_NowPlayer_Name.Text = ImportData[2, Now_Player];
                            lb_NowPlayer_Group.Text = ImportData[0, Now_Player];
                            lb_NowPlayer_Unit.Text = ImportData[1, Now_Player];
                            lb_NowPlayer_Match.Text = ImportData[3, Now_Player];
                        }
                    }
                }
                if (MsgBoxResult1 == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                {
                    //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                    //label1.Text = " 你選擇了按下”No“的按鈕！";
                    //MessageBox.Show("那你再考慮考慮");
                }
            }
            
        }

        /*private void CommandSend(int DeviceID, string OP_code)
        {
            serialPort1.Write(CommandConstruct(DeviceID, OP_code), 0, 16);
        }*/

        private byte[] CommandConstruct(int DeviceID, string OP_Code)//所有要發送出去的指令，都由這個函式組成，只要給 "機碼" 跟 "指令碼" 他會給你相對應的，你想要送出的 16Byte 的 Byte array (備註:當給的機碼為0則代表由PC發出的資訊的意思，而不用針對各個Device，而device收到由機碼 "FF" => 也就是ASCII "46 46" 時，不須回應RoutineAck)
        {
            if (DeviceID == 0)//若輸入的DeviceID=0，代表是PC發送的廣播，僅適用於OPCode="04"(開始評分),和OPCode="05"(結束評分)，輸入的DeviceID=0，就是機號要帶PC，PC機號是"FF"，"FF"ASCII是"46 46"
            {                
                //CommandConstruct/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                #region CommandConstruct
                byte[] Byte12_CommandConstruct = new byte[12] 
                { 
                    Convert.ToByte("0x21", 16),
                    Convert.ToByte("0x3B", 16),
                    Convert.ToByte("0x46", 16),
                    Convert.ToByte("0x46", 16),
                    Convert.ToByte("0x3"+OP_Code.Substring(0,1), 16),//Opcode
                    Convert.ToByte("0x3"+OP_Code.Substring(1,1), 16),//Opcode
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16)
                };
                string CRC_CommandConstruct = String.Format("{0,4:X}", CRC_any(Byte12_CommandConstruct, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
                CRC_CommandConstruct = CRC_CommandConstruct.Replace(" ", "0");
                byte[] Byte16_CommandConstruct = new byte[16] 
                { 
                    Convert.ToByte("0x21", 16),
                    Convert.ToByte("0x3B", 16),
                    Convert.ToByte("0x46", 16),
                    Convert.ToByte("0x46", 16),
                    Convert.ToByte("0x3"+OP_Code.Substring(0,1), 16),//Opcode
                    Convert.ToByte("0x3"+OP_Code.Substring(1,1), 16),//Opcode
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    //----------------------------
                    Convert.ToByte("0x00", 16),
                    Convert.ToByte("0x00", 16),
                    Convert.ToByte("0x"+CRC_CommandConstruct.Substring(0,2), 16),
                    Convert.ToByte("0x"+CRC_CommandConstruct.Substring(2,2), 16)
                };
                #endregion
                //CommandConstruct/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                return Byte16_CommandConstruct;
            }
            else//DeviceID != 0,代表是要跟各個Devie要數值，因此，機號的部分要帶,各個機號的名稱(ie,"01"~"07")
            {
                string DeviceName = "0"+DeviceID.ToString();//機號
                //CommandConstruct/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                #region CommandConstruct
                byte[] Byte12_CommandConstruct = new byte[12] 
                { 
                    Convert.ToByte("0x21", 16),
                    Convert.ToByte("0x3B", 16),
                    Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                    Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                    Convert.ToByte("0x3"+OP_Code.Substring(0,1), 16),//Opcode
                    Convert.ToByte("0x3"+OP_Code.Substring(1,1), 16),//Opcode
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16)
                };
                string CRC_CommandConstruct = String.Format("{0,4:X}", CRC_any(Byte12_CommandConstruct, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
                CRC_CommandConstruct = CRC_CommandConstruct.Replace(" ", "0");
                byte[] Byte16_CommandConstruct = new byte[16] 
                { 
                    Convert.ToByte("0x21", 16),
                    Convert.ToByte("0x3B", 16),
                    Convert.ToByte("0x3"+DeviceName.Substring(0, 1), 16),
                    Convert.ToByte("0x3"+DeviceName.Substring(1, 1), 16),
                    Convert.ToByte("0x3"+OP_Code.Substring(0,1), 16),//Opcode
                    Convert.ToByte("0x3"+OP_Code.Substring(1,1), 16),//Opcode
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    Convert.ToByte("0x30", 16),
                    //----------------------------
                    Convert.ToByte("0x00", 16),
                    Convert.ToByte("0x00", 16),
                    Convert.ToByte("0x"+CRC_CommandConstruct.Substring(0,2), 16),
                    Convert.ToByte("0x"+CRC_CommandConstruct.Substring(2,2), 16)
                };
                #endregion
                //CommandConstruct/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                return Byte16_CommandConstruct;
            }
            
        }

        //CRC////////////////////////////////////////////////////////////////////////////////
        #region CRC轉換
        static ulong CRC_any(byte[] blk_adr, ulong ulPoly, ulong ulInit, ulong ulXorOut, ulong ulMask)
        {
            ulong crc = ulInit;
            byte ucByte;
            int blk_len = blk_adr.Length;
            int i;
            bool iTopBitCRC;
            bool iTopBitByte;
            ulong ulTopBit;
            if (ulMask > 0xffff)
                ulTopBit = 0x80000000;
            else
                ulTopBit = ((ulMask + 1) >> 1);

            for (int j = 0; j < blk_len; j++)
            {
                ucByte = blk_adr[j];
                for (i = 0; i < 8; i++)
                {
                    iTopBitCRC = (crc & ulTopBit) != 0;
                    iTopBitByte = (ucByte & 0x80) != 0;
                    if (iTopBitCRC != iTopBitByte)
                    {
                        crc = (crc << 1) ^ ulPoly;
                    }
                    else
                    {
                        crc = (crc << 1);
                    }
                    ucByte <<= 1;
                }
            }
            return (ulong)((crc ^ ulXorOut) & ulMask);
        }
        #endregion
        //CRC/////////////////////////////////////////////////////////////////////////////////
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
        //ByteToHex///////////////////////////////////////////////////////////////////////////
        #region ByteToHex
        // --------------------------------------------------------------------------------
        //  @fn       ByteToHex()
        //  @brief    TODO
        //  @param    TODO
        // --------------------------------------------------------------------------------
        private string ByteToHex(byte[] comByte)
        {
            StringBuilder builder = new StringBuilder(comByte.Length * 3);
            foreach (byte data in comByte)
            {
                builder.Append(Convert.ToString(data, 16).PadLeft(2, '0').PadRight(3, ' '));
            }
            return builder.ToString().ToUpper();
        }
        #endregion
        //ByteToHex///////////////////////////////////////////////////////////////////////////
        //TextBox顯示/////////////////////////////////////////////////////////////////////////
        #region TextBox顯示
        delegate void SetTextbox1(string temp, System.Windows.Forms.TextBox tempTextbox);
        private void SetTextbox(string temp, System.Windows.Forms.TextBox tempTextbox)
        {
            if (tempTextbox.InvokeRequired)
            {
                SetTextbox1 d = new SetTextbox1(SetTextbox);
                this.Invoke(d, new object[] { temp, tempTextbox });
            }
            else
            {
                tempTextbox.Text = temp;
            }
        }
        #endregion

        private void btn_Abstain_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            DialogResult MsgBoxResult;//設置對話框的返回值
            MsgBoxResult = MessageBox.Show("該名選手確認棄權?",//對話框的顯示內容 
            "系統訊息",//對話框的標題 
            MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
            MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
            MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
            if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
            {
                //EasonFlow/////////////////////////////////////////////////////////
                btn_Demonstrate.Enabled = true;
                btn_Demonstrate.ForeColor = Color.Red;
                //EasonFlow/////////////////////////////////////////////////////////
                if (Select_Pose == "1")
                {
                    ImportData[16, Now_Player] = "1";
                    ImportData[6, Now_Player] = "-1";
                    ImportData[7, Now_Player] = "-1";
                    ImportData[8, Now_Player] = "-1";
                    ImportData[12, Now_Player] = "-1";
                    //原始分數///////////////////////////////////////////
                    ImportData[19, Now_Player] = "-1"; //第一品勢原始總分                   
                    ImportData[21, Now_Player] = "-1";//原始總分
                    //原始分數///////////////////////////////////////////
                }
                if (Select_Pose == "2")
                {
                    ImportData[16, Now_Player] = "1";
                    ImportData[6, Now_Player] = "-1";
                    ImportData[7, Now_Player] = "-1";
                    ImportData[8, Now_Player] = "-1";
                    ImportData[9, Now_Player] = "-1";
                    ImportData[10, Now_Player] = "-1";
                    ImportData[11, Now_Player] = "-1";
                    ImportData[12, Now_Player] = "-1";
                    //原始分數///////////////////////////////////////////
                    ImportData[19, Now_Player] = "-1"; //第一品勢原始總分                   
                    ImportData[20, Now_Player] = "-1"; //第二品勢原始總分   
                    ImportData[21, Now_Player] = "-1";//原始總分
                    //原始分數///////////////////////////////////////////
                }
                WriteExcelAll();


                if (Now_Player == Raw_CompetitorNumber)
                {
                    MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    Now_Player++;
                }

                lb_NowPlayer.Text = ImportData[15, Now_Player];
                lb_NowPlayer_Name.Text = ImportData[2, Now_Player];
                lb_NowPlayer_Group.Text = ImportData[0, Now_Player];
                lb_NowPlayer_Unit.Text = ImportData[1, Now_Player];
                lb_NowPlayer_Match.Text = ImportData[3, Now_Player];
            }
            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
            {

            }
        }

        //-------------------------------------------------------------------------------------------------
        String FilePath = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar;
        IniFile ini = new IniFile(System.Windows.Forms.Application.StartupPath + "\\FileName.ini");
        //-------------------------------------------------------------------------------------------------
        //連線
        string PortName = "";
        TaekwondoSerial.TaekwondoSerial TWS;
        private void btn_OPEN_COM_Click(object sender, EventArgs e)
        {
            if (btn_OPEN_COM.Text == "開啟")
            {
                try
                {
                    //if (TWS.serialOpen(cb_COM.Text))
                    //{
                        groupBox4.Enabled = false;
                        cb_COM.Enabled = false;
                        btn_OPEN_COM.Text = "斷線";
                        PortName = cb_COM.Text;
                        cb_COM.Text = "成功開啟";
                        btn_ImportSetting.Visible = true;
                        btn_ImportSetting.Enabled = true;
                        
                    //}
                    //else
                    //     throw new TimeoutException();
                }
                catch(Exception E)
                {
                    MessageBox.Show("連接阜開啟錯誤，請重新啟動程式，並確認連接阜", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
            }
            else
            {
                try
                {
                    if (MessageBox.Show("您確定要斷線嗎？", "Are you sure?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        disconnectDevice();
                    }
                }
                catch
                {
				    MessageBox.Show("連接埠錯誤，請重新啟動程式，並確認連接埠", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
				    this.Close();
                }
            }            
        }
        private void disconnectDevice()
        {
            try
            {
                if(TWS != null)
                    TWS.serialClose();
                btn_OPEN_COM.Text = "開啟";
                cb_COM.Enabled = true;
                cb_COM.Text = God.Yuan.GetCOMPorts()[0];
                btn_Demonstrate.Enabled = false;
                btn_Rating01.Enabled = false;
                btn_Rating02.Enabled = false;
                btn_Demonstrate02.Enabled = false;
                btn_Demonstrate01_Start.Enabled = false;
                btn_Demonstrate02_Start.Enabled = false;
            }
            catch
            {
                Console.WriteLine("disconnectDevice Error!", System.Reflection.MethodBase.GetCurrentMethod().Name);
            }
        }

        private void btn_Demonstrate01_Start_Click(object sender, EventArgs e)
        {
            RemoteControl_TimerStart = true;
            //EasonFlow//////////////////////////////////////////////////////////
            btn_Demonstrate01_Start.Enabled = false;
            btn_Demonstrate01_Start.ForeColor = Color.White;
            //----------------------------------------------
            btn_Rating01.Enabled = true;
            btn_Rating01.ForeColor = Color.Red;
            //EasonFlow//////////////////////////////////////////////////////////
        }

        private void btn_Demonstrate02_Start_Click(object sender, EventArgs e)
        {
            //EasonFlow///////////////////////////////////////////////////
            //把自己變白
            btn_Demonstrate02_Start.ForeColor = Color.White;
            btn_Demonstrate02_Start.Enabled = false;
            //把下一個變紅
            btn_Rating02.Enabled = true;
            btn_Rating02.ForeColor = Color.Red;
            //EasonFlow///////////////////////////////////////////////////
            RemoteControl_TimerStart = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(Environment.ExitCode);
            TWS.serialClose();
        }

        private void timer_CheckRatingState_Tick(object sender, EventArgs e)//程式開始時就一直執行
        {
            if (IsOpen_ImportSetting == true)//確認已經將資料載入了
            {
                if (ImportData[8, Now_Player] == "") 
                {
                    lb_Pose01_ScoreTotal.Text = "";
                }
                else if (ImportData[8, Now_Player] == "-1")
                {
                    lb_Pose01_ScoreTotal.Text = "棄權";
                }
                else 
                {
                    lb_Pose01_ScoreTotal.Text = ImportData[8, Now_Player];
                }

                //控制第2品勢按鍵顯示-------------------------------------------------------------------------
                if (lb_Pose01_ScoreTotal.Text != "" && lb_Pose02_ScoreTotal.Text == "" && Select_Pose == "2" && Now_Pose == 1)
                {
                    btn_Demonstrate02.ForeColor = Color.Red;
                    btn_Demonstrate02.Enabled = true;
                    //-------------------------------------------
                    btn_Demonstrate.ForeColor = Color.Black;
                    btn_Demonstrate.Enabled = false;
                }
                //控制第2品勢按鍵顯示-------------------------------------------------------------------------

                if (ImportData[11, Now_Player] == "")
                {
                    lb_Pose02_ScoreTotal.Text = "";                    
                }
                else if (ImportData[11, Now_Player] == "-1")
                {
                    lb_Pose02_ScoreTotal.Text = "棄權";
                }
                else
                {
                    lb_Pose02_ScoreTotal.Text = ImportData[11, Now_Player];
                }

                //lb_ScoreAverage.Text

                if (ImportData[12, Now_Player] == "")
                {
                    lb_ScoreAverage.Text = "";
                }
                else if (ImportData[12, Now_Player] == "-1")
                {
                    lb_ScoreAverage.Text = "棄權";
                }
                else
                {
                    lb_ScoreAverage.Text = ImportData[12, Now_Player];
                }

                //原始分數///////////////////////////////////////////////
                #region
                if (ImportData[19, Now_Player] == "")
                {
                    lb_Original_Score_SumAll_Pose1.Text = "";
                }
                else if (ImportData[19, Now_Player] == "-1")
                {
                    lb_Original_Score_SumAll_Pose1.Text = "棄權";
                }
                else
                {
                    lb_Original_Score_SumAll_Pose1.Text = ImportData[19, Now_Player];
                }

                if (ImportData[20, Now_Player] == "")
                {
                    lb_Original_Score_SumAll_Pose2.Text = "";
                }
                else if (ImportData[20, Now_Player] == "-1")
                {
                    lb_Original_Score_SumAll_Pose2.Text = "棄權";
                }
                else
                {
                    lb_Original_Score_SumAll_Pose2.Text = ImportData[20, Now_Player];
                }

                if (ImportData[21, Now_Player] == "")
                {
                    lb_Original_Score_SumAll_Pose1AddPose2.Text = "";
                }
                else if (ImportData[21, Now_Player] == "-1")
                {
                    lb_Original_Score_SumAll_Pose1AddPose2.Text = "棄權";
                }
                else
                {
                    lb_Original_Score_SumAll_Pose1AddPose2.Text = ImportData[21, Now_Player];
                }
                #endregion
                //原始分數///////////////////////////////////////////////

            }
        }

        private void btn_ReNew_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            DialogResult MsgBoxResult;//設置對話框的返回值
            MsgBoxResult = MessageBox.Show("確認重新評分?",//對話框的顯示內容 
            "系統訊息",//對話框的標題 
            MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
            MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
            MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
            if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
            {
                if (IsOpen_Rating == true && Select_Pose == "2")//當事2品勢時在關掉評分頁，因為2品勢會分數蓋過，但若只單一品勢不需
                {
                    
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////               
                    Rating.Close();
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                }
                //EasonFlow/////////////////////////////////////////////////////////
                btn_Demonstrate.Enabled = true;
                btn_Demonstrate.ForeColor = Color.Red;
                //------------------------------------------------
                btn_Demonstrate01_Start.ForeColor = Color.White;
                btn_Demonstrate01_Start.Enabled = false;
                btn_Rating01.ForeColor = Color.White;
                btn_Rating01.Enabled = false;
                btn_Demonstrate02.ForeColor = Color.White;
                btn_Demonstrate02.Enabled = false;
                btn_Demonstrate02_Start.ForeColor = Color.White;
                btn_Demonstrate02_Start.Enabled = false;
                btn_Rating02.ForeColor = Color.White;
                btn_Rating02.Enabled = false;
                //EasonFlow/////////////////////////////////////////////////////////
                if (Select_Pose == "1")
                {
                    btn_Demonstrate.Enabled = true;
                    /*
                    btn_Demonstrate01_Start.Enabled = true;
                    btn_Rating01.Enabled = true;
                    */
                }
                if (Select_Pose == "2")
                {
                    btn_Demonstrate.Enabled = true;
                    /*
                    btn_Demonstrate01_Start.Enabled = true;
                    btn_Rating01.Enabled = true;

                    btn_Demonstrate02.Enabled = true;
                    btn_Demonstrate02_Start.Enabled = true;
                    btn_Rating02.Enabled = true;
                    */
                }
                ImportData[6, Now_Player] = "";
                ImportData[7, Now_Player] = "";
                ImportData[8, Now_Player] = "";
                ImportData[9, Now_Player] = "";
                ImportData[10, Now_Player] = "";
                ImportData[11, Now_Player] = "";
                ImportData[12, Now_Player] = "";
                ImportData[13, Now_Player] = "";
                ImportData[14, Now_Player] = "";
                ImportData[16, Now_Player] = "";
                //原始分數/////////////////////////////////////////
                ImportData[19, Now_Player] = "";
                ImportData[20, Now_Player] = "";
                ImportData[21, Now_Player] = "";
                //原始分數/////////////////////////////////////////
                WriteExcelAll();
            }
            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
            {

            }


        }

        private void btn_ErrorFix_Click(object sender, EventArgs e)
        {
            //Set主控端評審================================================================================
            #region Set主控端評審
            Color Connect_Yes = Color.Green;//Color.Gray;
            if (Count_Referee == 1)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
            }
            if (Count_Referee == 2)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
                lb_device02.Text = "正\n常";
                lb_device02.ForeColor = Connect_Yes;
            }
            if (Count_Referee == 3)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
                lb_device02.Text = "正\n常";
                lb_device02.ForeColor = Connect_Yes;
                lb_device03.Text = "正\n常";
                lb_device03.ForeColor = Connect_Yes;
            }
            if (Count_Referee == 4)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
                lb_device02.Text = "正\n常";
                lb_device02.ForeColor = Connect_Yes;
                lb_device03.Text = "正\n常";
                lb_device03.ForeColor = Connect_Yes;
                lb_device04.Text = "正\n常";
                lb_device04.ForeColor = Connect_Yes;
            }
            if (Count_Referee == 5)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
                lb_device02.Text = "正\n常";
                lb_device02.ForeColor = Connect_Yes;
                lb_device03.Text = "正\n常";
                lb_device03.ForeColor = Connect_Yes;
                lb_device04.Text = "正\n常";
                lb_device04.ForeColor = Connect_Yes;
                lb_device05.Text = "正\n常";
                lb_device05.ForeColor = Connect_Yes;
            }
            if (Count_Referee == 6)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
                lb_device02.Text = "正\n常";
                lb_device02.ForeColor = Connect_Yes;
                lb_device03.Text = "正\n常";
                lb_device03.ForeColor = Connect_Yes;
                lb_device04.Text = "正\n常";
                lb_device04.ForeColor = Connect_Yes;
                lb_device05.Text = "正\n常";
                lb_device05.ForeColor = Connect_Yes;
                lb_device06.Text = "正\n常";
                lb_device06.ForeColor = Connect_Yes;
            }
            if (Count_Referee == 7)
            {
                lb_device01.Text = "正\n常";
                lb_device01.ForeColor = Connect_Yes;
                lb_device02.Text = "正\n常";
                lb_device02.ForeColor = Connect_Yes;
                lb_device03.Text = "正\n常";
                lb_device03.ForeColor = Connect_Yes;
                lb_device04.Text = "正\n常";
                lb_device04.ForeColor = Connect_Yes;
                lb_device05.Text = "正\n常";
                lb_device05.ForeColor = Connect_Yes;
                lb_device06.Text = "正\n常";
                lb_device06.ForeColor = Connect_Yes;
                lb_device07.Text = "正\n常";
                lb_device07.ForeColor = Connect_Yes;
            }
            #endregion
            //Set主控端評審================================================================================
            for (int i = 1; i <= Count_Referee; i++)
            {
                Flag_Correctness[i] = "0";
                Flag_Performance01[i] = "0";
                Flag_Performance02[i] = "0";
                Flag_Performance03[i] = "0";
                Score_Correctness[i] = "0";
                Score_Performance01[i] = "0";
                Score_Performance02[i] = "0";
                Score_Performance03[i] = "0";
                Flag_Pooling[i] = 0;
            }
            //斷線後將所有分數歸 0 
            btn_ErrorFix.Visible = false;//修復後，按鈕消失不見
           
        }


        //RankParameter////////////////////////////////////////////////////////////////////
        #region Rank Parameter
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
        //變數//////////////////////////////////////////////////////////////////////////////
        int Sum_ASCII_Number = 0;//編號的ascii相加
        double Final_Weighted_Score = 0;//最終加權分數
        //變數//////////////////////////////////////////////////////////////////////////////
        //排名//////////////////////////////////////////////
        //int RankPlace = 0;
        int SamePlace = 0;
        //排名//////////////////////////////////////////////
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
        #endregion
        //RankParameter////////////////////////////////////////////////////////////////////

        public void RankForAllPlayer()
        {
            //ini/////////////////////////////////////////////////////////////////////////////////////////////////////////
            ScoreRank_Total_Variable = new ScoreRank_Total[Now_Player + 1];
            for (int i = 0; i <= Now_Player; i++)
            {
                if (i == 0)
                {
                    ScoreRank_Total_Variable[0] = new ScoreRank_Total("Rank", "Number", "Name", -100000000000.0, 0, "");//因為後面有加權，要-10*10^12才夠小
                }
                else
                {
                    //算編號的ADCII總和///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    Sum_ASCII_Number = 0;

                    for (int j = 0; j < comt.Form1.ImportData[15, i].Length; j++)
                    {
                        string String_Substring = comt.Form1.ImportData[15, i].Substring(j, 1);
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

            Array.Sort(ScoreRank_Total_Variable, SortMethod_Total);
            int Counter_Quit = 0;

            //將名次填入陣列中/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            for (int k = Now_Player; k >= 1; k--)
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
                            comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                            //---------------------------
                            comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Now_Player - k + 1 - Counter_Quit - SamePlace) + "名";//國字名次
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
                            comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit).ToString();
                            ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit).ToString();
                            //---------------------------
                            comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Now_Player - k + 1 - Counter_Quit) + "名";//國字名次
                        }
                        //用原本陣列中的位置 = 給他名次 ， 從最高的名次給 到 最低
                    }
                    else
                    {
                        comt.Form1.ImportData[13, position] = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();//第一名                        
                        ScoreRank_Total_Variable[k].RealPlace = (Now_Player - k + 1 - Counter_Quit - SamePlace).ToString();
                        //---------------------------
                        comt.Form1.ImportData[14, position] = "第" + GetChineseNumber(Now_Player - k + 1 - Counter_Quit - SamePlace) + "名";//國字名次
                    }
                }
            }
            //將名次填入陣列中/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        }

        public static void WriteExcelExport()//匯出且排序
        {
            //引用Range類別
            Range myRange = null;

            //ExcelFileLocation
            _Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook myBook = myExcel.Workbooks.Open(ExcelFileLocation, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet mySheet = (Worksheet)myBook.Worksheets[1];

            myRange = (Range)mySheet.get_Range(myExcel.Cells[2, 1], myExcel.Cells[1 + Raw_CompetitorNumber, 22]);//原myRange = (Range)mySheet.get_Range(myExcel.Cells[2, 1], myExcel.Cells[1 + Raw_CompetitorNumber, 17]);

            myRange.Sort(
               mySheet.get_Range("J2", Type.Missing),
               Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
               Type.Missing,
               Type.Missing,
               Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
               Type.Missing,
               Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
               Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
               Type.Missing,
               Type.Missing,
               Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
               Microsoft.Office.Interop.Excel.XlSortMethod.xlStroke,
               Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortTextAsNumbers,
               Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
               Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal
           );

            myBook.Save();
            myBook.Close(false, Type.Missing, Type.Missing);//關閉Excel
            myExcel.Quit();//釋放Excel資源            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            myExcel = null;
            mySheet = null;
            //Range = null;
            myBook = null;
            GC.Collect();
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
        }

        private void btn_FTPUpload_Click(object sender, EventArgs e)
        {
            FTPUpload FTPUpload = new FTPUpload();
            FTPUpload.Show();
        }

        private void cb_COM_SelectedIndexChanged(object sender, EventArgs e)
        {
            btn_OPEN_COM.Enabled = true;
        }

        private void cb_COM_Click(object sender, EventArgs e)
        {
            cb_COM.Items.Clear();
            cb_COM.Items.AddRange(God.Yuan.GetCOMPorts());
            cb_COM.Text = God.Yuan.GetCOMPorts()[0];
        }
        bool isDeviceError = false;
        private void BGWorker_DeviceState_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (btn_OPEN_COM.Text == "斷線" && Count_Referee > 0)
                {
                    for (int i = 0; i < Count_Referee; i++)
                    {
                        try
                        {
                            DeviceState[i] = TWS.getReferee(i+1);
                        } catch (Exception E) {
                            DeviceState[i] = -1;
                          
                        }
                    }
                    isDeviceError = false;
                }
                else
                {
                    for (int i = 0; i < Count_Referee; i++)
                    {
                        DeviceState[i] = -1;
                    }
                }
                if(timer_Scoring.Enabled)
                {
                    for (int i = 0; i < Count_Referee; i++)
                    {
                        if (btn_OPEN_COM.Text == "斷線") {
                            if (DeviceState[i] != -1) {
                                for (int j = 0; j < Constant.ScoringItemNum; j++)
                                {
                                    if (!isFinished[j])
                                    { 
                                        float ScoreOrError =
                                            (j == (int)ScoringItem_4.Accuracy) ? Convert.ToSingle(TWS.GetRefereeGreen40Score_PK2(i + 1)) : //return value -1 means error, otherwise means socre
                                            (j == (int)ScoringItem_4.PowerSpeed) ? Convert.ToSingle(TWS.GetRefereeGreenPSScore_PK2(i + 1)) :
                                            (j == (int)ScoringItem_4.ForceTone) ? Convert.ToSingle(TWS.GetRefereeGreenFTScore_PK2(i + 1)) :
                                                                                    Convert.ToSingle(TWS.GetRefereeGreenSPScore_PK2(i + 1)); //TWS.GetRefereeGreen40Score_PK2
                                        
                                        if (ScoreOrError != -1)
                                        {
                                            ScoreArray[j][i+1] = ScoreOrError;
                                            RefereeStateDouble[i] |= 0x01 << j;
                                            if (j == 0) // correctness
                                            {
                                                Flag_Correctness[i+1] = "1";
                                            }
                                            if (j == 1) // Performance01
                                            {
                                                Flag_Performance01[i + 1] = "1";
                                            }
                                            if (j == 2) // Performance01
                                            {
                                                Flag_Performance02[i + 1] = "1";
                                            }
                                            if (j == 3) // Performance01
                                            {
                                                Flag_Performance03[i + 1] = "1";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception Ex) {
                Console.WriteLine("Device Error! " + Ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name);
                isDeviceError = true;
            }
        }

        private void BGWorker_DeviceState_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            for (int i = 0; i < Count_Referee; i++)
            {
                if (DeviceState[i] >= 0)
                {
                    Lb_Device[i].BackColor = Color.Lime;
                    Lb_Device[i].Text = "正\n常";
                }
                else
                {
                    Lb_Device[i].BackColor = Color.LightCoral;
                    Lb_Device[i].Text = "異\n常";
                }
                
                DisconnectCnt[i] = (DeviceState[i] >= 0) ? 0 : DisconnectCnt[i];
                if (btn_OPEN_COM.Text == "斷線" && Count_Referee > 0 && DeviceState[i] == -1)
                {
                    if (DisconnectCnt[i] >= 5)
                    {
                        if (DisconnectCnt[i] != 6)
                        {
                            timer_Device.Stop();
                            MessageBox.Show("第(" + (i + 1).ToString() + ")號 機台無回應", "錯誤訊息，請檢查裝置", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            DisconnectCnt[i] = 6;
                            timer_Device.Start(); //繼續檢查其他Device State
                        }
                    }
                    else
                    {
                        DisconnectCnt[i]++;
                    }   
                }
            }
            if (isDeviceError)
                disconnectDevice();
        }

        private void timer_Device_Tick(object sender, EventArgs e)
        {
            if (!BGWorker_DeviceState.IsBusy)
                BGWorker_DeviceState.RunWorkerAsync();
        }
        public System.Windows.Forms.Label[][] lb_score = new System.Windows.Forms.Label[Constant.ScoringItemNum][];
        private void timer_Scoring_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < Count_Referee; i++)
            {
                if (btn_OPEN_COM.Text == "斷線" && Count_Referee > 0)
                {
                    if (DeviceState[i] != -1)
                    {
                        for (int ItemCnt = 0; ItemCnt < Constant.ScoringItemNum; ItemCnt++)
                        {
                            if (!isFinished[ItemCnt]) {
                                if (((RefereeStateDouble[i] >> ItemCnt) & 0x01) == 1) {
                                    //lb_score[ItemCnt][i].Text = "OK";
                                }
                            }
                        }
                    }
                }
            }
            
            bool isAllFinished = true;
            for (int i = 0; i < isFinished.Length; i++)
            {
                if (!isFinished[i])
                {
                    isAllFinished = false;
                    break;
                }
            }
            if (isAllFinished) {
                //timer_Scoring.Stop();
                // Update UI
            }
        }
    }
}
