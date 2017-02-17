using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//--------------------------------------------
using System.IO;
using System.Net;
using System.Threading;


namespace comt
{
    public partial class FTPUpload : Form
    {

        //Excel檔案位置//////////////////////////////////////////////////////
        public string ExcelFileLocation;
        //Excel檔案位置//////////////////////////////////////////////////////
        string FileName = "";//只有檔名，不包含路徑
        string ServerAckMessage = "";

        //-------------------------------------------------------------------------------------------------
        String FilePath = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar;
        IniFile ini = new IniFile(System.Windows.Forms.Application.StartupPath + "\\FileName.ini");
        //-------------------------------------------------------------------------------------------------

        public FTPUpload()
        {
            InitializeComponent();
            ExcelFileLocation = comt.Form1.ExcelFileLocation;
            FileName = Path.GetFileName(ExcelFileLocation);
            tb_Folder.Text = ExcelFileLocation;
            //------------------------------------------------
            tb_IP.Text = ini.IniReadValue("Info", "IPAddress");
            //_displayLabel = lb_UploadState;
        }

        private void btn_Browse_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();         
            ExcelFileLocation = openFileDialog1.FileName;
            FileName = Path.GetFileName(ExcelFileLocation);
            tb_Folder.Text = ExcelFileLocation;
        }
        
        private Label _displayLabel;
        public Label DisplayLabel
        {
            get { return _displayLabel; }
            set { _displayLabel = value; }
        }
        /*
        public delegate void SetText(string text);
        private void SetLbText()
        {
            // 如果返回 True ，则访问控件的线程不是创建控件的线程
            if (btn_Upload.InvokeRequired)
            {
                // 实例一个委托，匿名方法，
                SetText st = new SetText(delegate(string text)
                {
                    // 改变 Label 的 Text
                    btn_Upload.Text = text;
                });

                // 把调用权交给创建控件的线程，带上参数
                btn_Upload.Invoke(st, "我是另一个线程---Invoke 方式");
            }
            else
            {
                btn_Upload.Text = "此控件是我创建的---Invoke 方式";
            }
        }
        */


        private void btn_Upload_Click(object sender, EventArgs e)
        {
            if (tb_IP.Text == "" || tb_Folder.Text == "")
            {
                MessageBox.Show("請確實填入IP位置 及 檔案位置", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                UpLoadStateDisplay UpLoadStateDisplay = new UpLoadStateDisplay();
                UpLoadStateDisplay.Show();
                #region mark
                /*
                DialogResult MsgBoxResult1;//設置對話框的返回值
                MsgBoxResult1 = MessageBox.Show("確認上傳檔案 < " + FileName + " >",//對話框的顯示內容 
                "系統訊息",//對話框的標題 
                MessageBoxButtons.YesNo,//定義對話框的按鈕，這里定義了YSE和NO兩個按鈕 
                MessageBoxIcon.Exclamation,//定義對話框內的圖表式樣，這里是一個黃色三角型內加一個感嘆號 
                MessageBoxDefaultButton.Button2);//定義對話框的按鈕式樣
                if (MsgBoxResult1 == DialogResult.Yes)//如果對話框的返回值是YES（按"Y"按鈕）
                {
                */
                    #region mark
                    /*
                    //判斷是否連上線///////////////////////////////////////////////////////////////////
                    #region 判斷是否連上線
                    bool Wifi_Enable = true;//判斷是否連上線
                    string strHostName = Dns.GetHostName();
                    // 取得本機的IpHostEntry類別實體，MSDN建議新的用法
                    IPHostEntry iphostentry = Dns.GetHostEntry(strHostName);
                    // 取得所有 IP 位址
                    foreach (IPAddress ipaddress in iphostentry.AddressList)
                    {
                        // 只取得IP V4的Address
                        if (ipaddress.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            string IPv4_PlaceFirst255 = ipaddress.ToString().Substring(0, ipaddress.ToString().IndexOf("."));
                            if (IPv4_PlaceFirst255 == "168" || IPv4_PlaceFirst255 == "0" || IPv4_PlaceFirst255 == "127")
                            {
                                Wifi_Enable = false;
                                MessageBox.Show("請檢察網路連線狀態", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    #endregion
                    //判斷是否連上線///////////////////////////////////////////////////////////////////
                    */
                    #endregion
                /*
                    if (Wifi_Enable == false)
                    {
                        textBox1.AppendText("上傳失敗 ( 無網路連線 )");
                        textBox1.AppendText(Environment.NewLine);
                    }
                    else
                    {*/
                #endregion
                ftp ftpClient = new ftp(@"ftp://" + tb_IP.Text + "/", "ftp", "ftp");
                //ftp(@"ServerIP", "使用者名稱", "使用者密碼");
                /* Upload a File */
                string IsUploadSuccess = "";
                IsUploadSuccess = ftpClient.upload(FileName, ExcelFileLocation);
                ServerAckMessage = IsUploadSuccess + " ( " + FileName + " )";
                textBox1.AppendText(ServerAckMessage);
                textBox1.AppendText(Environment.NewLine);
                    
                if (IsUploadSuccess == "上傳成功")
                {
                    UpLoadStateDisplay.Close();
                    DialogResult MsgBoxResult_Close;
                    MsgBoxResult_Close = MessageBox.Show(IsUploadSuccess + "，請問是否關閉程式?",
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
                        this.Close();
                    }
                }
                if (IsUploadSuccess == "上傳失敗")
                {
                    UpLoadStateDisplay.Close();
                    btn_Upload.Text = "上傳";
                    MessageBox.Show("上傳失敗，請檢查網路狀態、伺服器IP設定及確認檔案是否存在後，再次點選上傳。");
                }

                # region mark
                //lb_UploadState.Text = "完畢";
                    //}
                /*
                }
                if (MsgBoxResult1 == DialogResult.No)//如果對話框的返回值是NO（按"N"按鈕）
                {
                }
                */
                #endregion
            }

            /*
            this.lb_UploadState.Invoke(
                new MethodInvoker(
                delegate
                {
                    this.lb_UploadState.Text = "處理中，請稍後...";
                }
            )
            );
            */
            /*
            _displayLabel.Invoke(new EventHandler(delegate
            {
                _displayLabel.Text = "處理中，請稍後...";
            }));
            */
           
           
        }

        private void btn_IPConfirm_Click(object sender, EventArgs e)
        {
            tb_Folder.Enabled = true;
            btn_Browse.Enabled = true;
            textBox1.Enabled = true;
            btn_Upload.Enabled = true;
            //---------------------------------------
            ini.IniWriteValue("Info", "IPAddress", tb_IP.Text);
        }
    }
}
