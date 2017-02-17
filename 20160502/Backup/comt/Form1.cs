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
//20130519_2316
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
        public static string[] Flag_Performance;//看表現性性評分了沒
        public static string[] Score_Correctness;//填入正確性分數
        public static string[] Score_Performance;//填入表現性分數

        public static int[] Flag_Pooling;//用來檢查是否有Ack//若5次以上沒Ack則表示Device無回應        
        //現行狀態State//////////////////////////////////////////////////

        //mutex//////////////////////////////////////////////////////////
        bool Mutex_DeviceRoutineAck = true;
        //mutex//////////////////////////////////////////////////////////

        //視窗///////////////////////////////////////////////////////////
        Demonstrate Demonstrate;
        Rating Rating;
        bool IsOpen_Demonstrate = false;
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

        string msg = "";
        public Form1()
        {
            InitializeComponent();

            /*
            //移動視窗到延伸桌面
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.DesktopLocation = new System.Drawing.Point(2000, 0);
            */

            tb_COM.Text = ini.IniReadValue("Info", "Port");
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
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
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
                                /*
                                Convert.ToByte("0x"+string_sum.Substring(36,2), 16),//13
                                Convert.ToByte("0x"+string_sum.Substring(39,2), 16),//14
                                Convert.ToByte("0x"+string_sum.Substring(42,2), 16),//15
                                Convert.ToByte("0x"+string_sum.Substring(45,2), 16),//16
                                */
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

                                            while (Flag_Correctness[Now_Device] == "1" && Flag_Performance[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
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

                                                if (Flag_Correctness[Now_Device] == "0" || Flag_Performance[Now_Device] == "0")
                                                {
                                                    break;
                                                }
                                                if (Milestone_Now_Device == Now_Device)//代表全部找完
                                                {
                                                    timer_AskForData.Enabled = false;//全部都找完了~停止Timer再發送訊息
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

                                        while (Flag_Correctness[Now_Device] == "1" && Flag_Performance[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
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

                                            if (Flag_Correctness[Now_Device] == "0" || Flag_Performance[Now_Device] == "0")
                                            {
                                                break;
                                            }
                                            if (Milestone_Now_Device == Now_Device)//代表全部找完
                                            {
                                                timer_AskForData.Enabled = false;//全部都找完了~停止Timer再發送訊息
                                                break;
                                            }
                                        }
                                        #endregion
                                        //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    }
                                    /*
                                    Now_Device++;
                                    while (Flag_Correctness[Now_Device] == "1" && Flag_Performance[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
                                    {
                                        Now_Device++;
                                        if (Flag_Correctness[Now_Device] == "0" || Flag_Performance[Now_Device] == "0")//只要正當和表現還有一個沒上傳就問這台
                                        {
                                            break;
                                        }
                                        if (Milestone_Now_Device == Now_Device)//代表全部找完
                                        {
                                            timer_AskForData.Enabled = false;
                                            break;
                                        }
                                    }
                                    */



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

                                        while (Flag_Correctness[Now_Device] == "1" && Flag_Performance[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
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

                                            if (Flag_Correctness[Now_Device] == "0" || Flag_Performance[Now_Device] == "0")
                                            {
                                                break;
                                            }
                                            if (Milestone_Now_Device == Now_Device)//代表全部找完
                                            {
                                                timer_AskForData.Enabled = false;//全部都找完了~停止Timer再發送訊息
                                                break;
                                            }
                                        }
                                        #endregion
                                        //決定下一個Device要去polling誰//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    }

                                    /*
                                    Now_Device++;
                                    while (Flag_Correctness[Now_Device] == "1" && Flag_Performance[Now_Device] == "1")//若這個機台表現性和正當性都上傳了，就找下一台
                                    {
                                        Now_Device++;
                                        if (Flag_Correctness[Now_Device] == "0" || Flag_Performance[Now_Device] == "0")//只要正當和表現還有一個沒上傳就問這台
                                        {
                                            break;
                                        }
                                        if (Milestone_Now_Device == Now_Device)//代表全部找完
                                        {
                                            timer_AskForData.Enabled = false;
                                            break;
                                        }
                                    }
                                    */
                                }
                                /*
                                if (Receive_Instruction == "21 3B 30" + " 3" + Now_Device + " 30 31 30 30 30 30 30 30 00 00 " + CRC_PCAskForData.Substring(0, 2) + " " + CRC_PCAskForData.Substring(2, 2) + " " && Referee_IsRating == false)//收到Device跟我要Data//收到Device的AskForData
                                {
                                }
                                */
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
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //21 3B 30 31 30 31 30 30 30 30 30 30 00 00 20 7D
            //21 3B 30 32 30 31 30 30 30 30 30 30 00 00 0D 39
            //21 3B 30 33 30 31 30 30 30 30 30 30 00 00 E6 1A
            //21 3B 30 34 30 31 30 30 30 30 30 30 00 00 57 B1 
            //0x21 0x3B 0x30 0x31 0x30 0x31 0x30 0x30 0x30 0x30 0x30 0x30 0x00 0x00 0x20 0x7D
            //0x21 0x3B 0x30 0x32 0x30 0x31 0x30 0x30 0x30 0x30 0x30 0x30 0x00 0x00 0x0D 0x39
            //0x21 0x3B 0x30 0x33 0x30 0x31 0x30 0x30 0x30 0x30 0x30 0x30 0x00 0x00 0xE6 0x1A
            //0x21 0x3B 0x30 0x34 0x30 0x31 0x30 0x30 0x30 0x30 0x30 0x30 0x00 0x00 0x57 0xB1
            //AskForData21 3B 30 31 30 31 30 30 30 30 30 30 00 00 20 7D              
            serialPort1.Write(CommandConstruct(7, "01"), 0, 16);//serialPort1.Write(Byte16_PCAskForData, 0, 16);
            this.textBox1.AppendText(Environment.NewLine);
            this.textBox1.AppendText("TX : " + ByteToHex(CommandConstruct(7, "01")));
            
            //tb_score1.Text = ByteToHex(CommandConstruct(5, "02"));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void bt_scan_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= 5; i++)//連續發送5次開始以防沒有收到
            {
                Thread.Sleep(50);
                serialPort1.Write(CommandConstruct(0, "04"), 0, 16);
            }
            //serialPort1.Write(CommandConstruct(0, "04"), 0, 16); //serialPort1.Write(Byte16_StartRating, 0, 16);
        }

        private void bt_establish_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= 5; i++)//連續發送5次結束以防沒有收到
            {
                Thread.Sleep(50);
                serialPort1.Write(CommandConstruct(0, "05"), 0, 16);
            }
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
            btn_Demonstrate.Enabled = true;
            //btn_Rating01.Enabled = true;
            if (Select_Pose == "2")//如果選擇比兩個品勢則剛開始的時候，將第2品勢的按鈕先enable = false (就是讓使用者按不到)
            {
                btn_Demonstrate02.Enabled = false;
                btn_Rating02.Enabled = false;
                btn_Demonstrate02_Start.Enabled = false;
                

            }
            else//如果只選擇比單一品勢則隱藏第二品勢按鈕
            {
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
            }

            btn_Next.Enabled = true;
            btn_ReNew.Enabled = true;
            btn_Abstain.Enabled = true;
            //--------------------------------------------------------------------------------

            lb_NowPlayer.Text = ImportData[15, Now_Player];
            lb_NowPlayer_Name.Text = ImportData[2, Now_Player];
            lb_NowPlayer_Group.Text = ImportData[0, Now_Player];
            lb_NowPlayer_Unit.Text = ImportData[1, Now_Player];
            lb_NowPlayer_Match.Text = ImportData[3, Now_Player];

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
            btn_Rating01.Enabled = true;
            //當按下表演的時候，必須要能夠 讓 計時開始的按鍵 和 評分的按鍵 能按

            //讓評分的地方知道表演視窗已開啟，以防日後評分視窗去關掉一個未開啟的視窗//////////////////////////////////////////////////
            IsOpen_Demonstrate = true;
            //讓評分的地方知道表演視窗已開啟，以防日後評分視窗去關掉一個未開啟的視窗//////////////////////////////////////////////////
            
            Now_Pose = 1;

            //廣播各Device說，已經開始表演了，可以打開機器讓評審一邊看比賽一邊評分了(廣播5次，Device有收到就有收到沒收到就算了)/////////////
            for (int i = 1; i <= 5; i++)//連續發送5次開始以防沒有收到
            {
                Thread.Sleep(50);
                serialPort1.Write(CommandConstruct(0, "04"), 0, 16);
            }
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
        }

        private void btn_Demonstrate02_Click(object sender, EventArgs e)
        {

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
                btn_Rating02.Enabled = true;
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
                    for (int i = 1; i <= 5; i++)
                    {
                        Thread.Sleep(50);
                        serialPort1.Write(CommandConstruct(0, "04"), 0, 16);
                    }

                    //廣播各Device說，已經開始表演了，可以打開機器讓評審一邊看比賽一邊評分了(廣播5次，Device有收到就有收到沒收到就算了///////////// 

                    //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                    Demonstrate02 = new Demonstrate();
                    Demonstrate02.FormClosed += new FormClosedEventHandler(Demonstrate02_FormClosed);
                    Demonstrate02.Show();
                    //偵測子畫面關閉，父畫面做事情---------------------------------------------------------
                }
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

        private void btn_Rating01_Click(object sender, EventArgs e)
        {
            IsOpen_Rating = true;//評分視窗開啟
            btn_Demonstrate02.Enabled = true;
            if (IsOpen_Demonstrate == true)
            {
                //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                comt.Demonstrate.timer.Enabled = false;
                RemoteControl_TimerStart = false;
                Demonstrate.Close();
                //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
            }            
            //收值陣列初始化/////////////////////////////////////////////////////////////////////
            Flag_Correctness = new string[8];
            Flag_Performance = new string[8];
            Score_Correctness = new string[8];
            Score_Performance = new string[8];
            Flag_Pooling = new int[8];

            for (int i = 1; i <= Count_Referee; i++) 
            {
                Flag_Correctness[i] = "0";
                Flag_Performance[i] = "0";
                Score_Correctness[i] = "0";
                Score_Performance[i] = "0";
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

            timer_AskForData.Enabled = true;
        }

        private void btn_Rating02_Click(object sender, EventArgs e)
        {
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
                Flag_Performance = new string[8];
                Score_Correctness = new string[8];
                Score_Performance = new string[8];
                Flag_Pooling = new int[8];

                for (int i = 1; i <= Count_Referee; i++)
                {
                    Flag_Correctness[i] = "0";
                    Flag_Performance[i] = "0";
                    Score_Correctness[i] = "0";
                    Score_Performance[i] = "0";
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

                timer_AskForData.Enabled = true;
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
                        btn_Demonstrate02.Enabled = false;
                        btn_Demonstrate02_Start.Enabled = false;
                        btn_Rating01.Enabled = false;
                        btn_Rating02.Enabled = false;
                        if (Now_Player == Raw_CompetitorNumber)
                        {
                            //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            DialogResult MsgBoxResult;//設置對話框的返回值
                            MsgBoxResult = MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者, " + "確認匯出並關閉視窗?",//對話框的顯示內容 
                            "系統訊息",//對話框的標題 
                            MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                            MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                            MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                            if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                            {
                                WriteExcelExport();
                                this.Close();
                            }
                            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                            {
                                //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                                //label1.Text = " 你選擇了按下”No“的按鈕！";
                                //MessageBox.Show("那你再考慮考慮");
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
                        btn_Demonstrate02.Enabled = false;
                        btn_Demonstrate02_Start.Enabled = false;
                        btn_Rating01.Enabled = false;
                        btn_Rating02.Enabled = false;
                        if (Now_Player == Raw_CompetitorNumber)
                        {
                            //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            DialogResult MsgBoxResult;//設置對話框的返回值
                            MsgBoxResult = MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者, " + "確認匯出並關閉視窗?",//對話框的顯示內容 
                            "系統訊息",//對話框的標題 
                            MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                            MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                            MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                            if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                            {
                                WriteExcelExport();
                                this.Close();
                            }
                            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                            {
                                //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                                //label1.Text = " 你選擇了按下”No“的按鈕！";
                                //MessageBox.Show("那你再考慮考慮");
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

        private void CommandSend(int DeviceID, string OP_code)
        {
            serialPort1.Write(CommandConstruct(DeviceID, OP_code), 0, 16);
        }

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

        private void timer_AskForData_Tick(object sender, EventArgs e)
        {            
            if (Flag_Pooling[Now_Device] <= 10)
            {
                Now_Wait_Device = Now_Device;
                Thread.Sleep(100);
                Flag_Pooling[Now_Device]++;
                serialPort1.Write(CommandConstruct(Now_Device, "01"), 0, 16);                
            }
            else//五次以上無回應，跳出錯誤訊息 
            {
                timer_AskForData.Enabled = false;
                MessageBox.Show("第(" + Now_Device.ToString() + ")號 機台無回應", "錯誤訊息，請檢查裝置", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btn_ErrorFix.Visible = true;//一旦斷線，顯示修復按鈕，待所有機台就緒，按下修復按鈕
            }
            /*
            Score_Correctness;//填入正確性分數
            Score_Performance;//填入表現性分數
            */
        }

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
                if (Select_Pose == "1")
                {
                    ImportData[16, Now_Player] = "1";
                    ImportData[6, Now_Player] = "-1";
                    ImportData[7, Now_Player] = "-1";
                    ImportData[8, Now_Player] = "-1";
                    ImportData[12, Now_Player] = "-1";
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

        private void btn_OPEN_COM_Click(object sender, EventArgs e)
        {
            try 
            {
                ini.IniWriteValue("Info", "Port", tb_COM.Text);
                serialPort1.PortName = tb_COM.Text;
                serialPort1.Open();
                serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort1_DataReceived);
                groupBox4.Enabled = false;
                tb_COM.Text = "成功開啟";
                btn_ImportSetting.Enabled = true;
            }
            catch(Exception ex)
            {                
                MessageBox.Show("連接阜開啟錯誤，請重新啟動程式，並確認連接阜", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            
        }

        private void btn_Demonstrate01_Start_Click(object sender, EventArgs e)
        {
            RemoteControl_TimerStart = true;
        }

        private void btn_Demonstrate02_Start_Click(object sender, EventArgs e)
        {
            RemoteControl_TimerStart = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(Environment.ExitCode);  
            /*
            //this.Close();
            //MessageBox.Show("只有<" + Raw_CompetitorNumber + ">位參賽者", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            DialogResult MsgBoxResult;//設置對話框的返回值
            MsgBoxResult = MessageBox.Show("確認關閉視窗?",//對話框的顯示內容 
            "系統訊息",//對話框的標題 
            MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
            MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
            MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
            if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
            {                
            }
            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
            {
                //MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                //label1.Text = " 你選擇了按下”No“的按鈕！";
                //MessageBox.Show("那你再考慮考慮");
            }            
            */
            /*
            try
            {
                serialPort1.Close();
                if (IsOpen_Rating == true)
                {
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////               
                    Rating.Close();
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                }
                if (IsOpen_Demonstrate == true)
                {
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                    comt.Demonstrate.timer.Enabled = false;
                    RemoteControl_TimerStart = false;
                    Demonstrate.Close();
                    //進入評分畫面前把表演畫面計時停掉，並且關閉表演畫面/////////////////////////////////
                }
                timer_AskForData.Enabled = false;//將polling Device的timer關閉
                timer_CheckRatingState.Enabled = false;//關掉視窗要把Timer停掉
                
                //e.Cancel = true;
                comt.Rating.Result.Close();
                comt.Result.Rank.Close();
                this.Close();
                Environment.Exit(Environment.ExitCode);               
                System.Windows.Forms.Application.Exit();
            }
            catch(Exception ex)
            {

            }
            */
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
                if (Select_Pose == "1")
                {
                    btn_Demonstrate.Enabled = true;
                    btn_Demonstrate01_Start.Enabled = true;
                    btn_Rating01.Enabled = true;
                }
                if (Select_Pose == "2")
                {
                    btn_Demonstrate.Enabled = true;
                    btn_Demonstrate01_Start.Enabled = true;
                    btn_Rating01.Enabled = true;

                    btn_Demonstrate02.Enabled = true;
                    btn_Demonstrate02_Start.Enabled = true;
                    btn_Rating02.Enabled = true;
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
                WriteExcelAll();
            }
            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
            {

            }


        }

        private void btn_ErrorFix_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= Count_Referee; i++)
            {
                Flag_Correctness[i] = "0";
                Flag_Performance[i] = "0";
                Score_Correctness[i] = "0";
                Score_Performance[i] = "0";
                Flag_Pooling[i] = 0;
            }
            //斷線後將所有分數歸 0 
            btn_ErrorFix.Visible = false;//修復後，按鈕消失不見
            timer_AskForData.Enabled = true;//修復後啟動TIMER繼續POLLING
        }

        public static void WriteExcelExport()//匯出且排序
        {
            //引用Range類別
            Range myRange = null;

            //ExcelFileLocation
            _Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook myBook = myExcel.Workbooks.Open(ExcelFileLocation, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet mySheet = (Worksheet)myBook.Worksheets[1];

            myRange = (Range)mySheet.get_Range(myExcel.Cells[2, 1], myExcel.Cells[1 + Raw_CompetitorNumber, 17]);

            myRange.Sort(
               mySheet.get_Range("N2", Type.Missing),
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
            /*
            DialogResult MsgBoxResult;//設置對話框的返回值
            MsgBoxResult = MessageBox.Show("確認匯出?",//對話框的顯示內容 
            "確認刪除檔案",//對話框的標題 
            MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
            MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
            MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
            if (MsgBoxResult == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
            {
                MessageBox.Show("Yes", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (MsgBoxResult == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
            {
                MessageBox.Show("No", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //this.label1.ForeColor = System.Drawing.Color.Blue;//字體顏色設定
                //label1.Text = " 你選擇了按下”No“的按鈕！";
                //MessageBox.Show("那你再考慮考慮");
            }
            */
            //this.Close();
        }
        //TextBox顯示/////////////////////////////////////////////////////////////////////////
    }
}
