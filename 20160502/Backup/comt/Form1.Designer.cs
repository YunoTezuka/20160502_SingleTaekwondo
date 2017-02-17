namespace comt
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.bt_scan = new System.Windows.Forms.Button();
            this.bt_establish = new System.Windows.Forms.Button();
            this.btn_ImportSetting = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.btn_form2 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.tb_DeviceName = new System.Windows.Forms.TextBox();
            this.lb_Deviece_title = new System.Windows.Forms.Label();
            this.btn_DeviceNameCheck = new System.Windows.Forms.Button();
            this.btn_Demonstrate = new System.Windows.Forms.Button();
            this.btn_Rating01 = new System.Windows.Forms.Button();
            this.btn_Demonstrate02 = new System.Windows.Forms.Button();
            this.btn_Rating02 = new System.Windows.Forms.Button();
            this.btn_Next = new System.Windows.Forms.Button();
            this.btn_ReNew = new System.Windows.Forms.Button();
            this.btn_Abstain = new System.Windows.Forms.Button();
            this.timer_AskForData = new System.Windows.Forms.Timer(this.components);
            this.lb_NowPlayer_Title = new System.Windows.Forms.Label();
            this.lb_NowPlayer = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pnl_background03 = new System.Windows.Forms.Panel();
            this.pnl_background02 = new System.Windows.Forms.Panel();
            this.btn_Demonstrate02_Start = new System.Windows.Forms.Button();
            this.pnl_background01 = new System.Windows.Forms.Panel();
            this.btn_Demonstrate01_Start = new System.Windows.Forms.Button();
            this.pnl_background04 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_Export = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.pnl_ScoreAverage = new System.Windows.Forms.Panel();
            this.lb_ScoreAverage = new System.Windows.Forms.Label();
            this.lb_ScoreAverage_Title = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Name = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Group = new System.Windows.Forms.Label();
            this.pnl_Pose02_ScoreTotal = new System.Windows.Forms.Panel();
            this.lb_Pose02_ScoreTotal = new System.Windows.Forms.Label();
            this.pnl_Pose01_ScoreTotal = new System.Windows.Forms.Panel();
            this.lb_Pose01_ScoreTotal = new System.Windows.Forms.Label();
            this.lb_Pose02_ScoreTotal_Title = new System.Windows.Forms.Label();
            this.lb_Pose01_ScoreTotal_Title = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Match = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Match_Title = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Group_Title = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Unit = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Unit_Title = new System.Windows.Forms.Label();
            this.lb_NowPlayer_Name_Title = new System.Windows.Forms.Label();
            this.tb_COM = new System.Windows.Forms.TextBox();
            this.lb_COM_Setting_Title = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btn_OPEN_COM = new System.Windows.Forms.Button();
            this.timer_CheckRatingState = new System.Windows.Forms.Timer(this.components);
            this.btn_ErrorFix = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.pnl_background03.SuspendLayout();
            this.pnl_background02.SuspendLayout();
            this.pnl_background01.SuspendLayout();
            this.pnl_background04.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.pnl_ScoreAverage.SuspendLayout();
            this.pnl_Pose02_ScoreTotal.SuspendLayout();
            this.pnl_Pose01_ScoreTotal.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // serialPort1
            // 
            this.serialPort1.PortName = "COM95";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(869, 225);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(318, 208);
            this.textBox1.TabIndex = 0;
            this.textBox1.Visible = false;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(915, 164);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 22);
            this.textBox2.TabIndex = 1;
            this.textBox2.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(869, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "要資料";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(877, 167);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "傳送:";
            this.label1.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(831, 228);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "接收:";
            this.label2.Visible = false;
            // 
            // bt_scan
            // 
            this.bt_scan.Location = new System.Drawing.Point(869, 41);
            this.bt_scan.Name = "bt_scan";
            this.bt_scan.Size = new System.Drawing.Size(75, 23);
            this.bt_scan.TabIndex = 5;
            this.bt_scan.Text = "開始";
            this.bt_scan.UseVisualStyleBackColor = true;
            this.bt_scan.Visible = false;
            this.bt_scan.Click += new System.EventHandler(this.bt_scan_Click);
            // 
            // bt_establish
            // 
            this.bt_establish.Location = new System.Drawing.Point(869, 70);
            this.bt_establish.Name = "bt_establish";
            this.bt_establish.Size = new System.Drawing.Size(75, 23);
            this.bt_establish.TabIndex = 6;
            this.bt_establish.Text = "結束";
            this.bt_establish.UseVisualStyleBackColor = true;
            this.bt_establish.Visible = false;
            this.bt_establish.Click += new System.EventHandler(this.bt_establish_Click);
            // 
            // btn_ImportSetting
            // 
            this.btn_ImportSetting.Enabled = false;
            this.btn_ImportSetting.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_ImportSetting.Location = new System.Drawing.Point(405, 17);
            this.btn_ImportSetting.Name = "btn_ImportSetting";
            this.btn_ImportSetting.Size = new System.Drawing.Size(171, 58);
            this.btn_ImportSetting.TabIndex = 7;
            this.btn_ImportSetting.Text = "載入資料表";
            this.btn_ImportSetting.UseVisualStyleBackColor = true;
            this.btn_ImportSetting.Click += new System.EventHandler(this.ImportSetting_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(869, 128);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "Excel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btn_form2
            // 
            this.btn_form2.Location = new System.Drawing.Point(869, 99);
            this.btn_form2.Name = "btn_form2";
            this.btn_form2.Size = new System.Drawing.Size(75, 23);
            this.btn_form2.TabIndex = 9;
            this.btn_form2.Text = "RankOpen";
            this.btn_form2.UseVisualStyleBackColor = true;
            this.btn_form2.Visible = false;
            this.btn_form2.Click += new System.EventHandler(this.btn_form2_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(948, 13);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 13;
            this.button4.Text = "開啟Keyboard";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Visible = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // tb_DeviceName
            // 
            this.tb_DeviceName.Location = new System.Drawing.Point(948, 57);
            this.tb_DeviceName.Name = "tb_DeviceName";
            this.tb_DeviceName.Size = new System.Drawing.Size(62, 22);
            this.tb_DeviceName.TabIndex = 14;
            this.tb_DeviceName.Text = "05";
            this.tb_DeviceName.Visible = false;
            // 
            // lb_Deviece_title
            // 
            this.lb_Deviece_title.AutoSize = true;
            this.lb_Deviece_title.Location = new System.Drawing.Point(948, 42);
            this.lb_Deviece_title.Name = "lb_Deviece_title";
            this.lb_Deviece_title.Size = new System.Drawing.Size(32, 12);
            this.lb_Deviece_title.TabIndex = 15;
            this.lb_Deviece_title.Text = "機號:";
            this.lb_Deviece_title.Visible = false;
            // 
            // btn_DeviceNameCheck
            // 
            this.btn_DeviceNameCheck.Location = new System.Drawing.Point(1016, 57);
            this.btn_DeviceNameCheck.Name = "btn_DeviceNameCheck";
            this.btn_DeviceNameCheck.Size = new System.Drawing.Size(46, 23);
            this.btn_DeviceNameCheck.TabIndex = 16;
            this.btn_DeviceNameCheck.Text = "Done";
            this.btn_DeviceNameCheck.UseVisualStyleBackColor = true;
            this.btn_DeviceNameCheck.Visible = false;
            this.btn_DeviceNameCheck.Click += new System.EventHandler(this.btn_DeviceNameCheck_Click);
            // 
            // btn_Demonstrate
            // 
            this.btn_Demonstrate.Enabled = false;
            this.btn_Demonstrate.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Demonstrate.Location = new System.Drawing.Point(3, 6);
            this.btn_Demonstrate.Name = "btn_Demonstrate";
            this.btn_Demonstrate.Size = new System.Drawing.Size(114, 58);
            this.btn_Demonstrate.TabIndex = 17;
            this.btn_Demonstrate.Text = "第一品勢預備";
            this.btn_Demonstrate.UseVisualStyleBackColor = true;
            this.btn_Demonstrate.Click += new System.EventHandler(this.btn_Demonstrate_Click);
            // 
            // btn_Rating01
            // 
            this.btn_Rating01.Enabled = false;
            this.btn_Rating01.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Rating01.Location = new System.Drawing.Point(5, 6);
            this.btn_Rating01.Name = "btn_Rating01";
            this.btn_Rating01.Size = new System.Drawing.Size(105, 58);
            this.btn_Rating01.TabIndex = 18;
            this.btn_Rating01.Text = "第一品勢評分";
            this.btn_Rating01.UseVisualStyleBackColor = true;
            this.btn_Rating01.Click += new System.EventHandler(this.btn_Rating01_Click);
            // 
            // btn_Demonstrate02
            // 
            this.btn_Demonstrate02.Enabled = false;
            this.btn_Demonstrate02.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Demonstrate02.Location = new System.Drawing.Point(6, 5);
            this.btn_Demonstrate02.Name = "btn_Demonstrate02";
            this.btn_Demonstrate02.Size = new System.Drawing.Size(114, 58);
            this.btn_Demonstrate02.TabIndex = 19;
            this.btn_Demonstrate02.Text = "第二品勢預備";
            this.btn_Demonstrate02.UseVisualStyleBackColor = true;
            this.btn_Demonstrate02.Click += new System.EventHandler(this.btn_Demonstrate02_Click);
            // 
            // btn_Rating02
            // 
            this.btn_Rating02.Enabled = false;
            this.btn_Rating02.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Rating02.Location = new System.Drawing.Point(5, 5);
            this.btn_Rating02.Name = "btn_Rating02";
            this.btn_Rating02.Size = new System.Drawing.Size(106, 58);
            this.btn_Rating02.TabIndex = 20;
            this.btn_Rating02.Text = "第二品勢評分";
            this.btn_Rating02.UseVisualStyleBackColor = true;
            this.btn_Rating02.Click += new System.EventHandler(this.btn_Rating02_Click);
            // 
            // btn_Next
            // 
            this.btn_Next.Enabled = false;
            this.btn_Next.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Next.Location = new System.Drawing.Point(15, 78);
            this.btn_Next.Name = "btn_Next";
            this.btn_Next.Size = new System.Drawing.Size(346, 58);
            this.btn_Next.TabIndex = 21;
            this.btn_Next.Text = "下一位選手";
            this.btn_Next.UseVisualStyleBackColor = true;
            this.btn_Next.Click += new System.EventHandler(this.btn_Next_Click);
            // 
            // btn_ReNew
            // 
            this.btn_ReNew.Enabled = false;
            this.btn_ReNew.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_ReNew.Location = new System.Drawing.Point(15, 16);
            this.btn_ReNew.Name = "btn_ReNew";
            this.btn_ReNew.Size = new System.Drawing.Size(171, 58);
            this.btn_ReNew.TabIndex = 22;
            this.btn_ReNew.Text = "重新評分";
            this.btn_ReNew.UseVisualStyleBackColor = true;
            this.btn_ReNew.Click += new System.EventHandler(this.btn_ReNew_Click);
            // 
            // btn_Abstain
            // 
            this.btn_Abstain.Enabled = false;
            this.btn_Abstain.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Abstain.Location = new System.Drawing.Point(190, 16);
            this.btn_Abstain.Name = "btn_Abstain";
            this.btn_Abstain.Size = new System.Drawing.Size(171, 58);
            this.btn_Abstain.TabIndex = 23;
            this.btn_Abstain.Text = "棄權";
            this.btn_Abstain.UseVisualStyleBackColor = true;
            this.btn_Abstain.Click += new System.EventHandler(this.btn_Abstain_Click);
            // 
            // timer_AskForData
            // 
            this.timer_AskForData.Interval = 50;
            this.timer_AskForData.Tick += new System.EventHandler(this.timer_AskForData_Tick);
            // 
            // lb_NowPlayer_Title
            // 
            this.lb_NowPlayer_Title.AutoSize = true;
            this.lb_NowPlayer_Title.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Title.Location = new System.Drawing.Point(3, 18);
            this.lb_NowPlayer_Title.Name = "lb_NowPlayer_Title";
            this.lb_NowPlayer_Title.Size = new System.Drawing.Size(57, 20);
            this.lb_NowPlayer_Title.TabIndex = 24;
            this.lb_NowPlayer_Title.Text = "籤號:";
            // 
            // lb_NowPlayer
            // 
            this.lb_NowPlayer.AutoSize = true;
            this.lb_NowPlayer.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer.Location = new System.Drawing.Point(68, 18);
            this.lb_NowPlayer.Name = "lb_NowPlayer";
            this.lb_NowPlayer.Size = new System.Drawing.Size(30, 20);
            this.lb_NowPlayer.TabIndex = 25;
            this.lb_NowPlayer.Text = "---";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pnl_background03);
            this.groupBox1.Controls.Add(this.pnl_background02);
            this.groupBox1.Controls.Add(this.pnl_background01);
            this.groupBox1.Controls.Add(this.pnl_background04);
            this.groupBox1.Location = new System.Drawing.Point(399, 81);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(373, 174);
            this.groupBox1.TabIndex = 26;
            this.groupBox1.TabStop = false;
            // 
            // pnl_background03
            // 
            this.pnl_background03.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.pnl_background03.Controls.Add(this.btn_Rating01);
            this.pnl_background03.Location = new System.Drawing.Point(248, 15);
            this.pnl_background03.Name = "pnl_background03";
            this.pnl_background03.Size = new System.Drawing.Size(113, 69);
            this.pnl_background03.TabIndex = 34;
            // 
            // pnl_background02
            // 
            this.pnl_background02.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.pnl_background02.Controls.Add(this.btn_Demonstrate02_Start);
            this.pnl_background02.Controls.Add(this.btn_Demonstrate02);
            this.pnl_background02.Location = new System.Drawing.Point(9, 95);
            this.pnl_background02.Name = "pnl_background02";
            this.pnl_background02.Size = new System.Drawing.Size(233, 69);
            this.pnl_background02.TabIndex = 34;
            // 
            // btn_Demonstrate02_Start
            // 
            this.btn_Demonstrate02_Start.Enabled = false;
            this.btn_Demonstrate02_Start.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Demonstrate02_Start.Location = new System.Drawing.Point(123, 5);
            this.btn_Demonstrate02_Start.Name = "btn_Demonstrate02_Start";
            this.btn_Demonstrate02_Start.Size = new System.Drawing.Size(107, 58);
            this.btn_Demonstrate02_Start.TabIndex = 22;
            this.btn_Demonstrate02_Start.Text = "第二品勢開始";
            this.btn_Demonstrate02_Start.UseVisualStyleBackColor = true;
            this.btn_Demonstrate02_Start.Click += new System.EventHandler(this.btn_Demonstrate02_Start_Click);
            // 
            // pnl_background01
            // 
            this.pnl_background01.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.pnl_background01.Controls.Add(this.btn_Demonstrate01_Start);
            this.pnl_background01.Controls.Add(this.btn_Demonstrate);
            this.pnl_background01.Location = new System.Drawing.Point(10, 15);
            this.pnl_background01.Name = "pnl_background01";
            this.pnl_background01.Size = new System.Drawing.Size(233, 69);
            this.pnl_background01.TabIndex = 34;
            // 
            // btn_Demonstrate01_Start
            // 
            this.btn_Demonstrate01_Start.Enabled = false;
            this.btn_Demonstrate01_Start.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Demonstrate01_Start.Location = new System.Drawing.Point(123, 6);
            this.btn_Demonstrate01_Start.Name = "btn_Demonstrate01_Start";
            this.btn_Demonstrate01_Start.Size = new System.Drawing.Size(107, 58);
            this.btn_Demonstrate01_Start.TabIndex = 21;
            this.btn_Demonstrate01_Start.Text = "第一品勢開始";
            this.btn_Demonstrate01_Start.UseVisualStyleBackColor = true;
            this.btn_Demonstrate01_Start.Click += new System.EventHandler(this.btn_Demonstrate01_Start_Click);
            // 
            // pnl_background04
            // 
            this.pnl_background04.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.pnl_background04.Controls.Add(this.btn_Rating02);
            this.pnl_background04.Location = new System.Drawing.Point(248, 95);
            this.pnl_background04.Name = "pnl_background04";
            this.pnl_background04.Size = new System.Drawing.Size(113, 69);
            this.pnl_background04.TabIndex = 34;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_ReNew);
            this.groupBox2.Controls.Add(this.btn_Next);
            this.groupBox2.Controls.Add(this.btn_Abstain);
            this.groupBox2.Location = new System.Drawing.Point(399, 261);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(373, 144);
            this.groupBox2.TabIndex = 27;
            this.groupBox2.TabStop = false;
            // 
            // btn_Export
            // 
            this.btn_Export.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Export.Location = new System.Drawing.Point(950, 96);
            this.btn_Export.Name = "btn_Export";
            this.btn_Export.Size = new System.Drawing.Size(171, 58);
            this.btn_Export.TabIndex = 24;
            this.btn_Export.Text = "匯出";
            this.btn_Export.UseVisualStyleBackColor = true;
            this.btn_Export.Visible = false;
            this.btn_Export.Click += new System.EventHandler(this.btn_Export_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.pnl_ScoreAverage);
            this.groupBox3.Controls.Add(this.lb_ScoreAverage_Title);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Name);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Group);
            this.groupBox3.Controls.Add(this.pnl_Pose02_ScoreTotal);
            this.groupBox3.Controls.Add(this.pnl_Pose01_ScoreTotal);
            this.groupBox3.Controls.Add(this.lb_Pose02_ScoreTotal_Title);
            this.groupBox3.Controls.Add(this.lb_Pose01_ScoreTotal_Title);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Match);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Match_Title);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Group_Title);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Unit);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Unit_Title);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Name_Title);
            this.groupBox3.Controls.Add(this.lb_NowPlayer);
            this.groupBox3.Controls.Add(this.lb_NowPlayer_Title);
            this.groupBox3.Location = new System.Drawing.Point(23, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(362, 456);
            this.groupBox3.TabIndex = 28;
            this.groupBox3.TabStop = false;
            // 
            // pnl_ScoreAverage
            // 
            this.pnl_ScoreAverage.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.pnl_ScoreAverage.Controls.Add(this.lb_ScoreAverage);
            this.pnl_ScoreAverage.Location = new System.Drawing.Point(9, 346);
            this.pnl_ScoreAverage.Name = "pnl_ScoreAverage";
            this.pnl_ScoreAverage.Size = new System.Drawing.Size(347, 76);
            this.pnl_ScoreAverage.TabIndex = 42;
            // 
            // lb_ScoreAverage
            // 
            this.lb_ScoreAverage.AutoSize = true;
            this.lb_ScoreAverage.Font = new System.Drawing.Font("新細明體", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_ScoreAverage.ForeColor = System.Drawing.Color.Yellow;
            this.lb_ScoreAverage.Location = new System.Drawing.Point(141, 21);
            this.lb_ScoreAverage.Name = "lb_ScoreAverage";
            this.lb_ScoreAverage.Size = new System.Drawing.Size(59, 40);
            this.lb_ScoreAverage.TabIndex = 41;
            this.lb_ScoreAverage.Text = "---";
            // 
            // lb_ScoreAverage_Title
            // 
            this.lb_ScoreAverage_Title.AutoSize = true;
            this.lb_ScoreAverage_Title.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_ScoreAverage_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_ScoreAverage_Title.Location = new System.Drawing.Point(137, 317);
            this.lb_ScoreAverage_Title.Name = "lb_ScoreAverage_Title";
            this.lb_ScoreAverage_Title.Size = new System.Drawing.Size(85, 24);
            this.lb_ScoreAverage_Title.TabIndex = 40;
            this.lb_ScoreAverage_Title.Text = "總平均";
            // 
            // lb_NowPlayer_Name
            // 
            this.lb_NowPlayer_Name.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lb_NowPlayer_Name.AutoSize = true;
            this.lb_NowPlayer_Name.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Name.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Name.Location = new System.Drawing.Point(68, 43);
            this.lb_NowPlayer_Name.Name = "lb_NowPlayer_Name";
            this.lb_NowPlayer_Name.Size = new System.Drawing.Size(30, 20);
            this.lb_NowPlayer_Name.TabIndex = 27;
            this.lb_NowPlayer_Name.Text = "---";
            this.lb_NowPlayer_Name.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lb_NowPlayer_Group
            // 
            this.lb_NowPlayer_Group.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lb_NowPlayer_Group.AutoSize = true;
            this.lb_NowPlayer_Group.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Group.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Group.Location = new System.Drawing.Point(68, 69);
            this.lb_NowPlayer_Group.Name = "lb_NowPlayer_Group";
            this.lb_NowPlayer_Group.Size = new System.Drawing.Size(30, 20);
            this.lb_NowPlayer_Group.TabIndex = 31;
            this.lb_NowPlayer_Group.Text = "---";
            this.lb_NowPlayer_Group.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pnl_Pose02_ScoreTotal
            // 
            this.pnl_Pose02_ScoreTotal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.pnl_Pose02_ScoreTotal.Controls.Add(this.lb_Pose02_ScoreTotal);
            this.pnl_Pose02_ScoreTotal.Location = new System.Drawing.Point(189, 218);
            this.pnl_Pose02_ScoreTotal.Name = "pnl_Pose02_ScoreTotal";
            this.pnl_Pose02_ScoreTotal.Size = new System.Drawing.Size(167, 76);
            this.pnl_Pose02_ScoreTotal.TabIndex = 39;
            // 
            // lb_Pose02_ScoreTotal
            // 
            this.lb_Pose02_ScoreTotal.AutoSize = true;
            this.lb_Pose02_ScoreTotal.Font = new System.Drawing.Font("新細明體", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_Pose02_ScoreTotal.ForeColor = System.Drawing.Color.Yellow;
            this.lb_Pose02_ScoreTotal.Location = new System.Drawing.Point(34, 31);
            this.lb_Pose02_ScoreTotal.Name = "lb_Pose02_ScoreTotal";
            this.lb_Pose02_ScoreTotal.Size = new System.Drawing.Size(59, 40);
            this.lb_Pose02_ScoreTotal.TabIndex = 37;
            this.lb_Pose02_ScoreTotal.Text = "---";
            // 
            // pnl_Pose01_ScoreTotal
            // 
            this.pnl_Pose01_ScoreTotal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.pnl_Pose01_ScoreTotal.Controls.Add(this.lb_Pose01_ScoreTotal);
            this.pnl_Pose01_ScoreTotal.Location = new System.Drawing.Point(9, 218);
            this.pnl_Pose01_ScoreTotal.Name = "pnl_Pose01_ScoreTotal";
            this.pnl_Pose01_ScoreTotal.Size = new System.Drawing.Size(167, 76);
            this.pnl_Pose01_ScoreTotal.TabIndex = 38;
            // 
            // lb_Pose01_ScoreTotal
            // 
            this.lb_Pose01_ScoreTotal.AutoSize = true;
            this.lb_Pose01_ScoreTotal.Font = new System.Drawing.Font("新細明體", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_Pose01_ScoreTotal.ForeColor = System.Drawing.Color.Yellow;
            this.lb_Pose01_ScoreTotal.Location = new System.Drawing.Point(37, 33);
            this.lb_Pose01_ScoreTotal.Name = "lb_Pose01_ScoreTotal";
            this.lb_Pose01_ScoreTotal.Size = new System.Drawing.Size(59, 40);
            this.lb_Pose01_ScoreTotal.TabIndex = 36;
            this.lb_Pose01_ScoreTotal.Text = "---";
            // 
            // lb_Pose02_ScoreTotal_Title
            // 
            this.lb_Pose02_ScoreTotal_Title.AutoSize = true;
            this.lb_Pose02_ScoreTotal_Title.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_Pose02_ScoreTotal_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_Pose02_ScoreTotal_Title.Location = new System.Drawing.Point(189, 191);
            this.lb_Pose02_ScoreTotal_Title.Name = "lb_Pose02_ScoreTotal_Title";
            this.lb_Pose02_ScoreTotal_Title.Size = new System.Drawing.Size(160, 24);
            this.lb_Pose02_ScoreTotal_Title.TabIndex = 35;
            this.lb_Pose02_ScoreTotal_Title.Text = "第二品勢總分";
            // 
            // lb_Pose01_ScoreTotal_Title
            // 
            this.lb_Pose01_ScoreTotal_Title.AutoSize = true;
            this.lb_Pose01_ScoreTotal_Title.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_Pose01_ScoreTotal_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_Pose01_ScoreTotal_Title.Location = new System.Drawing.Point(9, 191);
            this.lb_Pose01_ScoreTotal_Title.Name = "lb_Pose01_ScoreTotal_Title";
            this.lb_Pose01_ScoreTotal_Title.Size = new System.Drawing.Size(160, 24);
            this.lb_Pose01_ScoreTotal_Title.TabIndex = 34;
            this.lb_Pose01_ScoreTotal_Title.Text = "第一品勢總分";
            // 
            // lb_NowPlayer_Match
            // 
            this.lb_NowPlayer_Match.AutoSize = true;
            this.lb_NowPlayer_Match.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Match.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Match.Location = new System.Drawing.Point(68, 128);
            this.lb_NowPlayer_Match.Name = "lb_NowPlayer_Match";
            this.lb_NowPlayer_Match.Size = new System.Drawing.Size(30, 20);
            this.lb_NowPlayer_Match.TabIndex = 33;
            this.lb_NowPlayer_Match.Text = "---";
            // 
            // lb_NowPlayer_Match_Title
            // 
            this.lb_NowPlayer_Match_Title.AutoSize = true;
            this.lb_NowPlayer_Match_Title.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Match_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Match_Title.Location = new System.Drawing.Point(3, 128);
            this.lb_NowPlayer_Match_Title.Name = "lb_NowPlayer_Match_Title";
            this.lb_NowPlayer_Match_Title.Size = new System.Drawing.Size(57, 20);
            this.lb_NowPlayer_Match_Title.TabIndex = 32;
            this.lb_NowPlayer_Match_Title.Text = "賽事:";
            // 
            // lb_NowPlayer_Group_Title
            // 
            this.lb_NowPlayer_Group_Title.AutoSize = true;
            this.lb_NowPlayer_Group_Title.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Group_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Group_Title.Location = new System.Drawing.Point(3, 69);
            this.lb_NowPlayer_Group_Title.Name = "lb_NowPlayer_Group_Title";
            this.lb_NowPlayer_Group_Title.Size = new System.Drawing.Size(57, 20);
            this.lb_NowPlayer_Group_Title.TabIndex = 30;
            this.lb_NowPlayer_Group_Title.Text = "組別:";
            // 
            // lb_NowPlayer_Unit
            // 
            this.lb_NowPlayer_Unit.AutoSize = true;
            this.lb_NowPlayer_Unit.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Unit.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Unit.Location = new System.Drawing.Point(68, 99);
            this.lb_NowPlayer_Unit.Name = "lb_NowPlayer_Unit";
            this.lb_NowPlayer_Unit.Size = new System.Drawing.Size(30, 20);
            this.lb_NowPlayer_Unit.TabIndex = 29;
            this.lb_NowPlayer_Unit.Text = "---";
            // 
            // lb_NowPlayer_Unit_Title
            // 
            this.lb_NowPlayer_Unit_Title.AutoSize = true;
            this.lb_NowPlayer_Unit_Title.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Unit_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Unit_Title.Location = new System.Drawing.Point(3, 99);
            this.lb_NowPlayer_Unit_Title.Name = "lb_NowPlayer_Unit_Title";
            this.lb_NowPlayer_Unit_Title.Size = new System.Drawing.Size(57, 20);
            this.lb_NowPlayer_Unit_Title.TabIndex = 28;
            this.lb_NowPlayer_Unit_Title.Text = "單位:";
            // 
            // lb_NowPlayer_Name_Title
            // 
            this.lb_NowPlayer_Name_Title.AutoSize = true;
            this.lb_NowPlayer_Name_Title.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_NowPlayer_Name_Title.ForeColor = System.Drawing.Color.Blue;
            this.lb_NowPlayer_Name_Title.Location = new System.Drawing.Point(3, 43);
            this.lb_NowPlayer_Name_Title.Name = "lb_NowPlayer_Name_Title";
            this.lb_NowPlayer_Name_Title.Size = new System.Drawing.Size(57, 20);
            this.lb_NowPlayer_Name_Title.TabIndex = 26;
            this.lb_NowPlayer_Name_Title.Text = "姓名:";
            // 
            // tb_COM
            // 
            this.tb_COM.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tb_COM.Location = new System.Drawing.Point(126, 14);
            this.tb_COM.Multiline = true;
            this.tb_COM.Name = "tb_COM";
            this.tb_COM.Size = new System.Drawing.Size(141, 35);
            this.tb_COM.TabIndex = 29;
            // 
            // lb_COM_Setting_Title
            // 
            this.lb_COM_Setting_Title.AutoSize = true;
            this.lb_COM_Setting_Title.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_COM_Setting_Title.Location = new System.Drawing.Point(6, 21);
            this.lb_COM_Setting_Title.Name = "lb_COM_Setting_Title";
            this.lb_COM_Setting_Title.Size = new System.Drawing.Size(114, 20);
            this.lb_COM_Setting_Title.TabIndex = 30;
            this.lb_COM_Setting_Title.Text = "連接阜設定:";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btn_OPEN_COM);
            this.groupBox4.Controls.Add(this.tb_COM);
            this.groupBox4.Controls.Add(this.lb_COM_Setting_Title);
            this.groupBox4.Location = new System.Drawing.Point(399, 410);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(373, 58);
            this.groupBox4.TabIndex = 31;
            this.groupBox4.TabStop = false;
            // 
            // btn_OPEN_COM
            // 
            this.btn_OPEN_COM.Font = new System.Drawing.Font("新細明體", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_OPEN_COM.Location = new System.Drawing.Point(280, 14);
            this.btn_OPEN_COM.Name = "btn_OPEN_COM";
            this.btn_OPEN_COM.Size = new System.Drawing.Size(81, 35);
            this.btn_OPEN_COM.TabIndex = 31;
            this.btn_OPEN_COM.Text = "開啟";
            this.btn_OPEN_COM.UseVisualStyleBackColor = true;
            this.btn_OPEN_COM.Click += new System.EventHandler(this.btn_OPEN_COM_Click);
            // 
            // timer_CheckRatingState
            // 
            this.timer_CheckRatingState.Tick += new System.EventHandler(this.timer_CheckRatingState_Tick);
            // 
            // btn_ErrorFix
            // 
            this.btn_ErrorFix.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btn_ErrorFix.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_ErrorFix.ForeColor = System.Drawing.Color.Red;
            this.btn_ErrorFix.Location = new System.Drawing.Point(589, 17);
            this.btn_ErrorFix.Name = "btn_ErrorFix";
            this.btn_ErrorFix.Size = new System.Drawing.Size(171, 58);
            this.btn_ErrorFix.TabIndex = 32;
            this.btn_ErrorFix.Text = "故障排除";
            this.btn_ErrorFix.UseVisualStyleBackColor = false;
            this.btn_ErrorFix.Visible = false;
            this.btn_ErrorFix.Click += new System.EventHandler(this.btn_ErrorFix_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(784, 473);
            this.Controls.Add(this.btn_Export);
            this.Controls.Add(this.btn_ErrorFix);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btn_DeviceNameCheck);
            this.Controls.Add(this.lb_Deviece_title);
            this.Controls.Add(this.tb_DeviceName);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.btn_form2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btn_ImportSetting);
            this.Controls.Add(this.bt_establish);
            this.Controls.Add(this.bt_scan);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "主控台";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.pnl_background03.ResumeLayout(false);
            this.pnl_background02.ResumeLayout(false);
            this.pnl_background01.ResumeLayout(false);
            this.pnl_background04.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.pnl_ScoreAverage.ResumeLayout(false);
            this.pnl_ScoreAverage.PerformLayout();
            this.pnl_Pose02_ScoreTotal.ResumeLayout(false);
            this.pnl_Pose02_ScoreTotal.PerformLayout();
            this.pnl_Pose01_ScoreTotal.ResumeLayout(false);
            this.pnl_Pose01_ScoreTotal.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button bt_scan;
        private System.Windows.Forms.Button bt_establish;
        private System.Windows.Forms.Button btn_ImportSetting;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btn_form2;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox tb_DeviceName;
        private System.Windows.Forms.Label lb_Deviece_title;
        private System.Windows.Forms.Button btn_DeviceNameCheck;
        private System.Windows.Forms.Button btn_Demonstrate;
        private System.Windows.Forms.Button btn_Rating01;
        private System.Windows.Forms.Button btn_Demonstrate02;
        private System.Windows.Forms.Button btn_Rating02;
        private System.Windows.Forms.Button btn_Next;
        private System.Windows.Forms.Button btn_ReNew;
        private System.Windows.Forms.Button btn_Abstain;
        private System.Windows.Forms.Timer timer_AskForData;
        private System.Windows.Forms.Label lb_NowPlayer_Title;
        private System.Windows.Forms.Label lb_NowPlayer;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lb_NowPlayer_Name;
        private System.Windows.Forms.Label lb_NowPlayer_Name_Title;
        private System.Windows.Forms.TextBox tb_COM;
        private System.Windows.Forms.Label lb_COM_Setting_Title;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btn_OPEN_COM;
        private System.Windows.Forms.Label lb_NowPlayer_Unit;
        private System.Windows.Forms.Label lb_NowPlayer_Unit_Title;
        private System.Windows.Forms.Label lb_NowPlayer_Group_Title;
        private System.Windows.Forms.Label lb_NowPlayer_Group;
        private System.Windows.Forms.Label lb_NowPlayer_Match_Title;
        private System.Windows.Forms.Label lb_NowPlayer_Match;
        private System.Windows.Forms.Button btn_Demonstrate01_Start;
        private System.Windows.Forms.Button btn_Demonstrate02_Start;
        private System.Windows.Forms.Panel pnl_background01;
        private System.Windows.Forms.Panel pnl_background02;
        private System.Windows.Forms.Panel pnl_background03;
        private System.Windows.Forms.Panel pnl_background04;
        private System.Windows.Forms.Timer timer_CheckRatingState;
        private System.Windows.Forms.Label lb_Pose02_ScoreTotal_Title;
        private System.Windows.Forms.Label lb_Pose01_ScoreTotal_Title;
        private System.Windows.Forms.Label lb_Pose02_ScoreTotal;
        private System.Windows.Forms.Label lb_Pose01_ScoreTotal;
        private System.Windows.Forms.Panel pnl_Pose02_ScoreTotal;
        private System.Windows.Forms.Panel pnl_Pose01_ScoreTotal;
        private System.Windows.Forms.Button btn_ErrorFix;
        private System.Windows.Forms.Button btn_Export;
        private System.Windows.Forms.Panel pnl_ScoreAverage;
        private System.Windows.Forms.Label lb_ScoreAverage;
        private System.Windows.Forms.Label lb_ScoreAverage_Title;
    }
}

