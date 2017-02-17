using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace comt
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            serialPort1.Open();
            serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort1_DataReceived);
            _displayTextBox = textBox1;
        }


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

        private TextBox _displayTextBox;
        public TextBox DisplayTextBox
        {
            get { return _displayTextBox; }
            set { _displayTextBox = value;}
        }

        string string_sum = "";
        string ReceiveRawData = "";
        [STAThread]
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            //Thread.Sleep(3);
            /*
            int ByteSize = serialPort1.BytesToRead;     //read一次時收到的data buffer size
            byte[] BufferData = new byte[ByteSize];
            */
            //////////////////////////////////////////////////////////
            int bytes = serialPort1.BytesToRead;
            // create a byte array to hold the awaiting data
            byte[] comBuffer = new byte[bytes];
            // read the data and store it
            serialPort1.Read(comBuffer, 0, bytes);
            //display the data to the user
            //DisplayData(MessageType.Incoming, ByteToHex(comBuffer) + "\n");
            this.textBox1.Invoke(
                new MethodInvoker(
                delegate
                {
                    this.textBox1.AppendText(ByteToHex(comBuffer)+"\n");
                    //this.textBox1.Text += " ";
                }
            )
            );
            /*
            _displayTextBox.Invoke(new EventHandler(delegate
            {
                ReceiveRawData = ByteToHex(comBuffer);
                _displayTextBox.AppendText(ReceiveRawData + "\n");
                //_displayTextBox.AppendText(Environment.NewLine);
                _displayTextBox.ScrollToCaret();

            }));
            */
            //////////////////////////////////////////////////////////

            //-------------------------------
            /*
            serialPort1.Read(BufferData, 0, BufferData.Length);
            if (BitConverter.ToString(BufferData) != "")
            {
                //MessageBox.Show(BitConverter.ToString(BufferData));
                //textBox1.Text = BitConverter.ToString(BufferData);
                
                //string_sum = string_sum + Environment.NewLine;
                string_sum = BitConverter.ToString(BufferData);
                
                _displayTextBox.Invoke(new EventHandler(delegate
                {
                    _displayTextBox.AppendText(string_sum);
                    _displayTextBox.ScrollToCaret();
                }));
              
                //SetTextbox(string_sum, textBox1);
            }
            */
            //textBox1.Text = BufferData[0].ToString();
        }

        #region 顯示TextBox
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




    }
}
