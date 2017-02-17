namespace comt
{
    partial class Keyboard
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lb_score_1 = new System.Windows.Forms.Label();
            this.lb_score_2 = new System.Windows.Forms.Label();
            this.btn_sub_03 = new System.Windows.Forms.Button();
            this.btn_sub_01 = new System.Windows.Forms.Button();
            this.lb_subtimes_1 = new System.Windows.Forms.Label();
            this.lb_subtimes_2 = new System.Windows.Forms.Label();
            this.btn_func_ok1 = new System.Windows.Forms.Button();
            this.btn_func_ok2 = new System.Windows.Forms.Button();
            this.btn_func_cancel = new System.Windows.Forms.Button();
            this.btn_num_1 = new System.Windows.Forms.Button();
            this.btn_num_2 = new System.Windows.Forms.Button();
            this.btn_num_3 = new System.Windows.Forms.Button();
            this.btn_num_4 = new System.Windows.Forms.Button();
            this.btn_num_5 = new System.Windows.Forms.Button();
            this.btn_num_6 = new System.Windows.Forms.Button();
            this.btn_num_7 = new System.Windows.Forms.Button();
            this.btn_num_8 = new System.Windows.Forms.Button();
            this.btn_num_9 = new System.Windows.Forms.Button();
            this.btn_num_buff1 = new System.Windows.Forms.Button();
            this.btn_num_0 = new System.Windows.Forms.Button();
            this.btn_num_buff2 = new System.Windows.Forms.Button();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.tb_score_1 = new System.Windows.Forms.TextBox();
            this.tb_score_2 = new System.Windows.Forms.TextBox();
            this.timer_DataReady = new System.Windows.Forms.Timer(this.components);
            this.lb_Device_Name = new System.Windows.Forms.Label();
            this.tb_DeviceName = new System.Windows.Forms.TextBox();
            this.btn_DeviceNameCheck = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lb_score_1
            // 
            this.lb_score_1.AutoSize = true;
            this.lb_score_1.Location = new System.Drawing.Point(84, 216);
            this.lb_score_1.Name = "lb_score_1";
            this.lb_score_1.Size = new System.Drawing.Size(35, 12);
            this.lb_score_1.TabIndex = 0;
            this.lb_score_1.Text = "分數1";
            // 
            // lb_score_2
            // 
            this.lb_score_2.AutoSize = true;
            this.lb_score_2.Location = new System.Drawing.Point(181, 216);
            this.lb_score_2.Name = "lb_score_2";
            this.lb_score_2.Size = new System.Drawing.Size(35, 12);
            this.lb_score_2.TabIndex = 1;
            this.lb_score_2.Text = "分數2";
            // 
            // btn_sub_03
            // 
            this.btn_sub_03.Location = new System.Drawing.Point(3, 228);
            this.btn_sub_03.Name = "btn_sub_03";
            this.btn_sub_03.Size = new System.Drawing.Size(75, 23);
            this.btn_sub_03.TabIndex = 2;
            this.btn_sub_03.Text = "-0.3";
            this.btn_sub_03.UseVisualStyleBackColor = true;
            // 
            // btn_sub_01
            // 
            this.btn_sub_01.Location = new System.Drawing.Point(296, 228);
            this.btn_sub_01.Name = "btn_sub_01";
            this.btn_sub_01.Size = new System.Drawing.Size(75, 23);
            this.btn_sub_01.TabIndex = 3;
            this.btn_sub_01.Text = "-0.1";
            this.btn_sub_01.UseVisualStyleBackColor = true;
            // 
            // lb_subtimes_1
            // 
            this.lb_subtimes_1.AutoSize = true;
            this.lb_subtimes_1.Location = new System.Drawing.Point(130, 239);
            this.lb_subtimes_1.Name = "lb_subtimes_1";
            this.lb_subtimes_1.Size = new System.Drawing.Size(35, 12);
            this.lb_subtimes_1.TabIndex = 4;
            this.lb_subtimes_1.Text = "次數1";
            // 
            // lb_subtimes_2
            // 
            this.lb_subtimes_2.AutoSize = true;
            this.lb_subtimes_2.Location = new System.Drawing.Point(200, 239);
            this.lb_subtimes_2.Name = "lb_subtimes_2";
            this.lb_subtimes_2.Size = new System.Drawing.Size(35, 12);
            this.lb_subtimes_2.TabIndex = 5;
            this.lb_subtimes_2.Text = "次數2";
            // 
            // btn_func_ok1
            // 
            this.btn_func_ok1.Location = new System.Drawing.Point(66, 269);
            this.btn_func_ok1.Name = "btn_func_ok1";
            this.btn_func_ok1.Size = new System.Drawing.Size(75, 23);
            this.btn_func_ok1.TabIndex = 6;
            this.btn_func_ok1.Text = "確認1";
            this.btn_func_ok1.UseVisualStyleBackColor = true;
            this.btn_func_ok1.Click += new System.EventHandler(this.btn_func_ok1_Click);
            // 
            // btn_func_ok2
            // 
            this.btn_func_ok2.Location = new System.Drawing.Point(148, 269);
            this.btn_func_ok2.Name = "btn_func_ok2";
            this.btn_func_ok2.Size = new System.Drawing.Size(75, 23);
            this.btn_func_ok2.TabIndex = 7;
            this.btn_func_ok2.Text = "確認2";
            this.btn_func_ok2.UseVisualStyleBackColor = true;
            this.btn_func_ok2.Click += new System.EventHandler(this.btn_func_ok2_Click);
            // 
            // btn_func_cancel
            // 
            this.btn_func_cancel.Location = new System.Drawing.Point(230, 269);
            this.btn_func_cancel.Name = "btn_func_cancel";
            this.btn_func_cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_func_cancel.TabIndex = 8;
            this.btn_func_cancel.Text = "取消";
            this.btn_func_cancel.UseVisualStyleBackColor = true;
            // 
            // btn_num_1
            // 
            this.btn_num_1.Location = new System.Drawing.Point(66, 300);
            this.btn_num_1.Name = "btn_num_1";
            this.btn_num_1.Size = new System.Drawing.Size(75, 23);
            this.btn_num_1.TabIndex = 9;
            this.btn_num_1.Text = "1";
            this.btn_num_1.UseVisualStyleBackColor = true;
            // 
            // btn_num_2
            // 
            this.btn_num_2.Location = new System.Drawing.Point(149, 298);
            this.btn_num_2.Name = "btn_num_2";
            this.btn_num_2.Size = new System.Drawing.Size(75, 23);
            this.btn_num_2.TabIndex = 10;
            this.btn_num_2.Text = "2";
            this.btn_num_2.UseVisualStyleBackColor = true;
            // 
            // btn_num_3
            // 
            this.btn_num_3.Location = new System.Drawing.Point(230, 298);
            this.btn_num_3.Name = "btn_num_3";
            this.btn_num_3.Size = new System.Drawing.Size(75, 23);
            this.btn_num_3.TabIndex = 11;
            this.btn_num_3.Text = "3";
            this.btn_num_3.UseVisualStyleBackColor = true;
            // 
            // btn_num_4
            // 
            this.btn_num_4.Location = new System.Drawing.Point(67, 327);
            this.btn_num_4.Name = "btn_num_4";
            this.btn_num_4.Size = new System.Drawing.Size(75, 23);
            this.btn_num_4.TabIndex = 12;
            this.btn_num_4.Text = "4";
            this.btn_num_4.UseVisualStyleBackColor = true;
            // 
            // btn_num_5
            // 
            this.btn_num_5.Location = new System.Drawing.Point(148, 327);
            this.btn_num_5.Name = "btn_num_5";
            this.btn_num_5.Size = new System.Drawing.Size(75, 23);
            this.btn_num_5.TabIndex = 13;
            this.btn_num_5.Text = "5";
            this.btn_num_5.UseVisualStyleBackColor = true;
            // 
            // btn_num_6
            // 
            this.btn_num_6.Location = new System.Drawing.Point(229, 328);
            this.btn_num_6.Name = "btn_num_6";
            this.btn_num_6.Size = new System.Drawing.Size(75, 23);
            this.btn_num_6.TabIndex = 14;
            this.btn_num_6.Text = "6";
            this.btn_num_6.UseVisualStyleBackColor = true;
            // 
            // btn_num_7
            // 
            this.btn_num_7.Location = new System.Drawing.Point(67, 356);
            this.btn_num_7.Name = "btn_num_7";
            this.btn_num_7.Size = new System.Drawing.Size(75, 23);
            this.btn_num_7.TabIndex = 15;
            this.btn_num_7.Text = "7";
            this.btn_num_7.UseVisualStyleBackColor = true;
            // 
            // btn_num_8
            // 
            this.btn_num_8.Location = new System.Drawing.Point(149, 356);
            this.btn_num_8.Name = "btn_num_8";
            this.btn_num_8.Size = new System.Drawing.Size(75, 23);
            this.btn_num_8.TabIndex = 16;
            this.btn_num_8.Text = "8";
            this.btn_num_8.UseVisualStyleBackColor = true;
            // 
            // btn_num_9
            // 
            this.btn_num_9.Location = new System.Drawing.Point(230, 356);
            this.btn_num_9.Name = "btn_num_9";
            this.btn_num_9.Size = new System.Drawing.Size(75, 23);
            this.btn_num_9.TabIndex = 17;
            this.btn_num_9.Text = "9";
            this.btn_num_9.UseVisualStyleBackColor = true;
            // 
            // btn_num_buff1
            // 
            this.btn_num_buff1.Location = new System.Drawing.Point(66, 385);
            this.btn_num_buff1.Name = "btn_num_buff1";
            this.btn_num_buff1.Size = new System.Drawing.Size(75, 23);
            this.btn_num_buff1.TabIndex = 18;
            this.btn_num_buff1.Text = "預留1";
            this.btn_num_buff1.UseVisualStyleBackColor = true;
            // 
            // btn_num_0
            // 
            this.btn_num_0.Location = new System.Drawing.Point(148, 385);
            this.btn_num_0.Name = "btn_num_0";
            this.btn_num_0.Size = new System.Drawing.Size(75, 23);
            this.btn_num_0.TabIndex = 19;
            this.btn_num_0.Text = "0";
            this.btn_num_0.UseVisualStyleBackColor = true;
            // 
            // btn_num_buff2
            // 
            this.btn_num_buff2.Location = new System.Drawing.Point(230, 385);
            this.btn_num_buff2.Name = "btn_num_buff2";
            this.btn_num_buff2.Size = new System.Drawing.Size(75, 23);
            this.btn_num_buff2.TabIndex = 20;
            this.btn_num_buff2.Text = "保留2";
            this.btn_num_buff2.UseVisualStyleBackColor = true;
            // 
            // serialPort1
            // 
            this.serialPort1.PortName = "COM93";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(3, 2);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(368, 170);
            this.textBox1.TabIndex = 21;
            // 
            // tb_score_1
            // 
            this.tb_score_1.Location = new System.Drawing.Point(125, 214);
            this.tb_score_1.Name = "tb_score_1";
            this.tb_score_1.Size = new System.Drawing.Size(50, 22);
            this.tb_score_1.TabIndex = 22;
            // 
            // tb_score_2
            // 
            this.tb_score_2.Location = new System.Drawing.Point(222, 214);
            this.tb_score_2.Name = "tb_score_2";
            this.tb_score_2.Size = new System.Drawing.Size(50, 22);
            this.tb_score_2.TabIndex = 23;
            // 
            // timer_DataReady
            // 
            this.timer_DataReady.Interval = 1;
            // 
            // lb_Device_Name
            // 
            this.lb_Device_Name.AutoSize = true;
            this.lb_Device_Name.Location = new System.Drawing.Point(65, 184);
            this.lb_Device_Name.Name = "lb_Device_Name";
            this.lb_Device_Name.Size = new System.Drawing.Size(67, 12);
            this.lb_Device_Name.TabIndex = 24;
            this.lb_Device_Name.Text = "DeviceName:";
            // 
            // tb_DeviceName
            // 
            this.tb_DeviceName.Location = new System.Drawing.Point(132, 178);
            this.tb_DeviceName.Name = "tb_DeviceName";
            this.tb_DeviceName.Size = new System.Drawing.Size(100, 22);
            this.tb_DeviceName.TabIndex = 25;
            this.tb_DeviceName.Text = "05";
            // 
            // btn_DeviceNameCheck
            // 
            this.btn_DeviceNameCheck.Location = new System.Drawing.Point(238, 178);
            this.btn_DeviceNameCheck.Name = "btn_DeviceNameCheck";
            this.btn_DeviceNameCheck.Size = new System.Drawing.Size(45, 23);
            this.btn_DeviceNameCheck.TabIndex = 26;
            this.btn_DeviceNameCheck.Text = "確認";
            this.btn_DeviceNameCheck.UseVisualStyleBackColor = true;
            this.btn_DeviceNameCheck.Click += new System.EventHandler(this.btn_DeviceNameCheck_Click);
            // 
            // Keyboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 411);
            this.Controls.Add(this.btn_DeviceNameCheck);
            this.Controls.Add(this.tb_DeviceName);
            this.Controls.Add(this.lb_Device_Name);
            this.Controls.Add(this.tb_score_2);
            this.Controls.Add(this.tb_score_1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btn_num_buff2);
            this.Controls.Add(this.btn_num_0);
            this.Controls.Add(this.btn_num_buff1);
            this.Controls.Add(this.btn_num_9);
            this.Controls.Add(this.btn_num_8);
            this.Controls.Add(this.btn_num_7);
            this.Controls.Add(this.btn_num_6);
            this.Controls.Add(this.btn_num_5);
            this.Controls.Add(this.btn_num_4);
            this.Controls.Add(this.btn_num_3);
            this.Controls.Add(this.btn_num_2);
            this.Controls.Add(this.btn_num_1);
            this.Controls.Add(this.btn_func_cancel);
            this.Controls.Add(this.btn_func_ok2);
            this.Controls.Add(this.btn_func_ok1);
            this.Controls.Add(this.lb_subtimes_2);
            this.Controls.Add(this.lb_subtimes_1);
            this.Controls.Add(this.btn_sub_01);
            this.Controls.Add(this.btn_sub_03);
            this.Controls.Add(this.lb_score_2);
            this.Controls.Add(this.lb_score_1);
            this.Name = "Keyboard";
            this.Text = "Keyboard";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lb_score_1;
        private System.Windows.Forms.Label lb_score_2;
        private System.Windows.Forms.Button btn_sub_03;
        private System.Windows.Forms.Button btn_sub_01;
        private System.Windows.Forms.Label lb_subtimes_1;
        private System.Windows.Forms.Label lb_subtimes_2;
        private System.Windows.Forms.Button btn_func_ok1;
        private System.Windows.Forms.Button btn_func_ok2;
        private System.Windows.Forms.Button btn_func_cancel;
        private System.Windows.Forms.Button btn_num_1;
        private System.Windows.Forms.Button btn_num_2;
        private System.Windows.Forms.Button btn_num_3;
        private System.Windows.Forms.Button btn_num_4;
        private System.Windows.Forms.Button btn_num_5;
        private System.Windows.Forms.Button btn_num_6;
        private System.Windows.Forms.Button btn_num_7;
        private System.Windows.Forms.Button btn_num_8;
        private System.Windows.Forms.Button btn_num_9;
        private System.Windows.Forms.Button btn_num_buff1;
        private System.Windows.Forms.Button btn_num_0;
        private System.Windows.Forms.Button btn_num_buff2;
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox tb_score_1;
        private System.Windows.Forms.TextBox tb_score_2;
        private System.Windows.Forms.Timer timer_DataReady;
        private System.Windows.Forms.Label lb_Device_Name;
        private System.Windows.Forms.TextBox tb_DeviceName;
        private System.Windows.Forms.Button btn_DeviceNameCheck;
    }
}