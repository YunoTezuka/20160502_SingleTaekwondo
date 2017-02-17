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
    public partial class Keyboard : Form
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
        //-----------------------------------
        string Receive_Instruction;
        string string_sum;
        //21 3B 30 31 30 32 30 30 30 30 30 30
        //-------------------------------------------------------
        bool Referee_IsRating = false;
        bool Referee_IsRating02 = false;
        //-------------------------------------------------------
        byte[] Command_Score_Array = new byte[16];
        byte[] Command_Score_Array02 = new byte[16];
        //-------------------------------------------------------
        
        //-------------------------------------------------------
        string DeviceName;
        byte[] RoutineAck_Array;
        byte[] PCAskForData_Array;

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

        string Recrive_Data;
        public Keyboard()
        {
            InitializeComponent();
            serialPort1.Open();
            serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort1_DataReceived);
        }

        private TextBox _displayTextBox;
        public TextBox DisplayTextBox
        {
            get { return _displayTextBox; }
            set { _displayTextBox = value; }
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
        [STAThread]
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
            string_sum = string_sum + ByteToHex(comBuffer);//蒐集到的指令

            while (string_sum.Length >= 48)
            {
                firstCharacter = string_sum.IndexOf("21 3B ");
                if (firstCharacter != -1)
                {
                    Real_Instruction = string_sum.Substring(firstCharacter, 48);//集合成一指令
                    string_sum = SubString(string_sum, firstCharacter + 48, (string_sum.Length - firstCharacter - 48));//扣除掉該指令，剩下的指令

                    if (Real_Instruction.Length == 48 && Real_Instruction.Substring(0, 5) == "21 3B")//先檢查資料長度有無到&&標頭檔對不對
                    {
                        #region If內的
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

                        if (string.Format("{0,2:X}", Real_Instruction.Substring(42, 2)) + string.Format("{0,2:X}", Real_Instruction.Substring(45, 2)) == String_CRC16)//再檢查CRC對不對
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

                            if (Receive_Instruction == "21 3B 3" + DeviceName.Substring(0, 1) + " 3" + DeviceName.Substring(1, 1) + " 30 31 30 30 30 30 30 30 00 00 " + CRC_PCAskForData.Substring(0, 2) + " " + CRC_PCAskForData.Substring(2, 2) + " " && Referee_IsRating == false)//收到Device跟我要Data//收到Device的AskForData
                            {
                                this.textBox1.Invoke(
                                new MethodInvoker(
                                delegate
                                {
                                    this.textBox1.AppendText("PCAskForData" + Environment.NewLine);
                                }
                                )
                                );
                                serialPort1.Write(Byte16_RoutineAck, 0, 16);//回答說我還沒準備好但我有ACK給你代表裁判分數尚未輸入完成
                                this.textBox1.Invoke(
                                new MethodInvoker(
                                delegate
                                {
                                    this.textBox1.AppendText("TX : " + ByteToHex(Byte16_RoutineAck) + "Routine Ack" + Environment.NewLine);
                                }
                                )
                                );
                            }
                            else if (Receive_Instruction == "21 3B 3" + DeviceName.Substring(0, 1) + " 3" + DeviceName.Substring(1, 1) + " 30 31 30 30 30 30 30 30 00 00 " + CRC_PCAskForData.Substring(0, 2) + " " + CRC_PCAskForData.Substring(2, 2) + " " && Referee_IsRating == true)
                            {
                                if (Referee_IsRating02 == false)
                                {
                                    this.textBox1.Invoke(
                                    new MethodInvoker(
                                    delegate
                                    {
                                        this.textBox1.AppendText("PCAskForData" + Environment.NewLine);
                                    }
                                    )
                                    );
                                    serialPort1.Write(Command_Score_Array, 0, 16);
                                    this.textBox1.Invoke(
                                    new MethodInvoker(
                                    delegate
                                    {
                                        this.textBox1.AppendText("TX : " + ByteToHex(Command_Score_Array) + "SendScore" + Environment.NewLine);
                                    }
                                    )
                                    );
                                }
                                else
                                {
                                    this.textBox1.Invoke(
                                    new MethodInvoker(
                                    delegate
                                    {
                                        this.textBox1.AppendText("PCAskForData" + Environment.NewLine);
                                    }
                                    )
                                    );
                                    serialPort1.Write(Command_Score_Array02, 0, 16);
                                    this.textBox1.Invoke(
                                    new MethodInvoker(
                                    delegate
                                    {
                                        this.textBox1.AppendText("TX : " + ByteToHex(Command_Score_Array02) + "SendScore" + Environment.NewLine);
                                    }
                                    )
                                    );
                                }
                            }
                            else if (Receive_Instruction == "21 3B 46 46 30 34 30 30 30 30 30 30 00 00 " + CRC_StartRating.Substring(0, 2) + " " + CRC_StartRating.Substring(2, 2) + " ")//收到開始評分指令
                            {

                                this.textBox1.Invoke(
                                new MethodInvoker(
                                delegate
                                {
                                    this.textBox1.AppendText("StartRating" + Environment.NewLine);
                                }
                                )
                                );

                                serialPort1.Write(Byte16_RoutineAck, 0, 16);
                                this.textBox1.Invoke(
                                new MethodInvoker(
                                delegate
                                {
                                    this.textBox1.AppendText("TX : " + ByteToHex(Byte16_RoutineAck) + "Routine Ack" + Environment.NewLine);
                                }
                                )
                                );

                            }
                            else if (Receive_Instruction == "21 3B 46 46 30 35 30 30 30 30 30 30 00 00 " + CRC_EndRating.Substring(0, 2) + " " + CRC_EndRating.Substring(2, 2) + " ")//收到結束評分指令
                            {

                                this.textBox1.Invoke(
                                new MethodInvoker(
                                delegate
                                {
                                    this.textBox1.AppendText("EndRating" + Environment.NewLine);
                                }
                                )
                                );

                                serialPort1.Write(Byte16_RoutineAck, 0, 16);

                                this.textBox1.Invoke(
                                new MethodInvoker(
                                delegate
                                {
                                    this.textBox1.AppendText("TX : " + ByteToHex(Byte16_RoutineAck) + "Routine Ack" + Environment.NewLine);
                                }
                                )
                                );

                            }
                            //serialPort1.Write(RoutineAck_Array, 0, 16);
                        }
                        else//如果CRC錯誤，放棄該筆指令，將收指令的string清空，再讓他繼續去收
                        {
                            //string_sum = "";
                        }
                        #endregion
                    }
                }

            }
            
            

            
            

        }

        private void btn_func_ok1_Click(object sender, EventArgs e)
        {   
            //正確性-----------------------------------------------------------
            //Byte
            //[12  ]  [3 4 ]   [5 6 ]   [7 8  ] [9 10 ] [11 12]  [13 14 15 16 ]
            //[!;  ]  [_ _ ]   [= = ]   [_ _  ] [_ _  ] [_ _  ]  [= = = =     ]
            //[標頭]  [機號]   [指令]   [資料1] [資料2] [資料3]  [CRC16       ]

            double score_double_1, score_double_2;

            score_double_1 = Convert.ToDouble(tb_score_1.Text);
            //score_double_2 = Convert.ToDouble(tb_score_2.Text);            

            Header_Byte1 = Convert.ToString(ASC("!"), 16);
            Header_Byte2 = Convert.ToString(ASC(";"), 16);
            Device_Byte1 = Convert.ToString(ASC(DeviceName.Substring(0, 1)), 16);
            Device_Byte2 = Convert.ToString(ASC(DeviceName.Substring(1, 1)), 16);
            Opcode_Byte1 = Convert.ToString(ASC("0"), 16);
            Opcode_Byte2 = Convert.ToString(ASC("3"), 16);

            Data1_Byte1 = Convert.ToString(ASC(((score_double_1 * 10) / 10).ToString()), 16);//十位數轉ASCII
            Data1_Byte2 = Convert.ToString(ASC(((score_double_1 * 10) % 10).ToString()), 16);//個位數轉ASCII

            Data2_Byte1 = "30";//Convert.ToString(ASC(((score_double_2 * 10) / 10).ToString()), 16);//十位數轉ASCII
            Data2_Byte2 = "30";//Convert.ToString(ASC(((score_double_2 * 10) % 10).ToString()), 16);//個位數轉ASCII

            Data3_Byte1 = "30";
            Data3_Byte2 = "30";

            byte[] CRC_Byte_Array = new byte[12] 
            { 
                Convert.ToByte("0x"+Header_Byte1, 16),
                Convert.ToByte("0x"+Header_Byte2, 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x"+Opcode_Byte1, 16),
                Convert.ToByte("0x"+Opcode_Byte2, 16),
                Convert.ToByte("0x"+Data1_Byte1, 16),
                Convert.ToByte("0x"+Data1_Byte2, 16),                
                Convert.ToByte("0x"+Data2_Byte1, 16),
                Convert.ToByte("0x"+Data2_Byte2, 16),                
                Convert.ToByte("0x"+Data3_Byte1, 16),
                Convert.ToByte("0x"+Data3_Byte2, 16),
            };            

            string String_CRC16;
            String_CRC16 = String.Format("{0,4:X}", CRC_any(CRC_Byte_Array, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
            String_CRC16 = String_CRC16.Replace(" ", "0"); 

            CRC_Byte1 = "00";
            CRC_Byte2 = "00";
            CRC_Byte3 = String_CRC16.Substring(0, 2);
            CRC_Byte4 = String_CRC16.Substring(2, 2);

            Command_Score_Array = new byte[16] 
            { 
                Convert.ToByte("0x"+Header_Byte1, 16),
                Convert.ToByte("0x"+Header_Byte2, 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x"+"30", 16),
                Convert.ToByte("0x"+"33", 16),
                Convert.ToByte("0x"+Data1_Byte1, 16),
                Convert.ToByte("0x"+Data1_Byte2, 16),                
                Convert.ToByte("0x"+Data2_Byte1, 16),
                Convert.ToByte("0x"+Data2_Byte2, 16),                
                Convert.ToByte("0x"+Data3_Byte1, 16),
                Convert.ToByte("0x"+Data3_Byte2, 16),
                Convert.ToByte("0x"+CRC_Byte1, 16),
                Convert.ToByte("0x"+CRC_Byte2, 16),
                Convert.ToByte("0x"+CRC_Byte3, 16),
                Convert.ToByte("0x"+CRC_Byte4, 16)
            };
            Referee_IsRating = true;//如果裁判已經評完分數了，就設為已評完分數，等PC下次問時再把分數丟出去
        }

        private void btn_func_ok2_Click(object sender, EventArgs e)
        {   
            //表現性-----------------------------------------------------------
            //Byte
            //[12  ]  [3 4 ]   [5 6 ]   [7 8  ] [9 10 ] [11 12]  [13 14 15 16 ]
            //[!;  ]  [_ _ ]   [= = ]   [_ _  ] [_ _  ] [_ _  ]  [= = = =     ]
            //[標頭]  [機號]   [指令]   [資料1] [資料2] [資料3]  [CRC16       ]

            double score_double_1, score_double_2;

            //score_double_1 = Convert.ToDouble(tb_score_1.Text);
            score_double_2 = Convert.ToDouble(tb_score_2.Text);

            Header_Byte1 = Convert.ToString(ASC("!"), 16);
            Header_Byte2 = Convert.ToString(ASC(";"), 16);
            Device_Byte1 = Convert.ToString(ASC(DeviceName.Substring(0, 1)), 16);
            Device_Byte2 = Convert.ToString(ASC(DeviceName.Substring(1, 1)), 16);
            Opcode_Byte1 = Convert.ToString(ASC("0"), 16);
            Opcode_Byte2 = Convert.ToString(ASC("6"), 16);

            Data1_Byte1 = "30";//Convert.ToString(ASC(((score_double_1 * 10) / 10).ToString()), 16);//十位數轉ASCII
            Data1_Byte2 = "30";//Convert.ToString(ASC(((score_double_1 * 10) % 10).ToString()), 16);//個位數轉ASCII

            Data2_Byte1 = Convert.ToString(ASC(((score_double_2 * 10) / 10).ToString()), 16);//十位數轉ASCII
            Data2_Byte2 = Convert.ToString(ASC(((score_double_2 * 10) % 10).ToString()), 16);//個位數轉ASCII

            Data3_Byte1 = "30";
            Data3_Byte2 = "30";

            byte[] CRC_Byte_Array = new byte[12] 
            { 
                Convert.ToByte("0x"+Header_Byte1, 16),
                Convert.ToByte("0x"+Header_Byte2, 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x"+Opcode_Byte1, 16),
                Convert.ToByte("0x"+Opcode_Byte2, 16),
                Convert.ToByte("0x"+Data1_Byte1, 16),
                Convert.ToByte("0x"+Data1_Byte2, 16),                
                Convert.ToByte("0x"+Data2_Byte1, 16),
                Convert.ToByte("0x"+Data2_Byte2, 16),                
                Convert.ToByte("0x"+Data3_Byte1, 16),
                Convert.ToByte("0x"+Data3_Byte2, 16),
            };

            string String_CRC16;
            String_CRC16 = String.Format("{0,4:X}", CRC_any(CRC_Byte_Array, 0x1021, 0xFFFF, 0x0000, 0xFFFF));
            String_CRC16 = String_CRC16.Replace(" ", "0");
            CRC_Byte1 = "00";
            CRC_Byte2 = "00";
            CRC_Byte3 = String_CRC16.Substring(0, 2);
            CRC_Byte4 = String_CRC16.Substring(2, 2);

            Command_Score_Array02 = new byte[16] 
            { 
                Convert.ToByte("0x"+Header_Byte1, 16),
                Convert.ToByte("0x"+Header_Byte2, 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(0, 1), 16),
                Convert.ToByte("0x"+"3"+DeviceName.Substring(1, 1), 16),
                Convert.ToByte("0x"+"30", 16),
                Convert.ToByte("0x"+"36", 16),
                Convert.ToByte("0x"+Data1_Byte1, 16),
                Convert.ToByte("0x"+Data1_Byte2, 16),                
                Convert.ToByte("0x"+Data2_Byte1, 16),
                Convert.ToByte("0x"+Data2_Byte2, 16),                
                Convert.ToByte("0x"+Data3_Byte1, 16),
                Convert.ToByte("0x"+Data3_Byte2, 16),
                Convert.ToByte("0x"+CRC_Byte1, 16),
                Convert.ToByte("0x"+CRC_Byte2, 16),
                Convert.ToByte("0x"+CRC_Byte3, 16),
                Convert.ToByte("0x"+CRC_Byte4, 16)
            };
            Referee_IsRating02 = true;//如果裁判已經評完分數了，就設為已評完分數，等PC下次問時再把分數丟出去
        }
        
        private void btn_DeviceNameCheck_Click(object sender, EventArgs e)
        {
            DeviceName = tb_DeviceName.Text;
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


            //StartRating///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region StartRating
            Byte12_StartRating = new byte[12] 
            { 
                Convert.ToByte("0x21", 16),
                Convert.ToByte("0x3B", 16),
                Convert.ToByte("0x46", 16),
                Convert.ToByte("0x46", 16),
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
                Convert.ToByte("0x46", 16),
                Convert.ToByte("0x46", 16),
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
                Convert.ToByte("0x46", 16),
                Convert.ToByte("0x46", 16),
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
                Convert.ToByte("0x46", 16),
                Convert.ToByte("0x46", 16),
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


        //CRC/////////////////////////////////////////////////////////////////////////////////
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
        delegate void SetTextbox1(string temp, TextBox tempTextbox);
        private void SetTextbox(string temp, TextBox tempTextbox)
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
        //TextBox顯示/////////////////////////////////////////////////////////////////////////

    }
}
