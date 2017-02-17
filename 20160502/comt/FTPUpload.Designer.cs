namespace comt
{
    partial class FTPUpload
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tb_IP = new System.Windows.Forms.TextBox();
            this.tb_Folder = new System.Windows.Forms.TextBox();
            this.lb_IP = new System.Windows.Forms.Label();
            this.lb_Folder = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btn_Upload = new System.Windows.Forms.Button();
            this.btn_Browse = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.lb_UploadState = new System.Windows.Forms.Label();
            this.btn_IPConfirm = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tb_IP
            // 
            this.tb_IP.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tb_IP.Location = new System.Drawing.Point(92, 7);
            this.tb_IP.Name = "tb_IP";
            this.tb_IP.Size = new System.Drawing.Size(226, 29);
            this.tb_IP.TabIndex = 0;
            // 
            // tb_Folder
            // 
            this.tb_Folder.Enabled = false;
            this.tb_Folder.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tb_Folder.Location = new System.Drawing.Point(92, 42);
            this.tb_Folder.Name = "tb_Folder";
            this.tb_Folder.Size = new System.Drawing.Size(226, 29);
            this.tb_Folder.TabIndex = 1;
            // 
            // lb_IP
            // 
            this.lb_IP.AutoSize = true;
            this.lb_IP.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_IP.Location = new System.Drawing.Point(12, 12);
            this.lb_IP.Name = "lb_IP";
            this.lb_IP.Size = new System.Drawing.Size(61, 21);
            this.lb_IP.TabIndex = 2;
            this.lb_IP.Text = "IP 位置";
            // 
            // lb_Folder
            // 
            this.lb_Folder.AutoSize = true;
            this.lb_Folder.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_Folder.Location = new System.Drawing.Point(12, 45);
            this.lb_Folder.Name = "lb_Folder";
            this.lb_Folder.Size = new System.Drawing.Size(74, 21);
            this.lb_Folder.TabIndex = 3;
            this.lb_Folder.Text = "檔案位置";
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textBox1.Location = new System.Drawing.Point(403, 187);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(383, 122);
            this.textBox1.TabIndex = 4;
            this.textBox1.Visible = false;
            // 
            // btn_Upload
            // 
            this.btn_Upload.Enabled = false;
            this.btn_Upload.Location = new System.Drawing.Point(153, 77);
            this.btn_Upload.Name = "btn_Upload";
            this.btn_Upload.Size = new System.Drawing.Size(130, 28);
            this.btn_Upload.TabIndex = 5;
            this.btn_Upload.Text = "上傳";
            this.btn_Upload.UseVisualStyleBackColor = true;
            this.btn_Upload.Click += new System.EventHandler(this.btn_Upload_Click);
            // 
            // btn_Browse
            // 
            this.btn_Browse.Enabled = false;
            this.btn_Browse.Location = new System.Drawing.Point(324, 43);
            this.btn_Browse.Name = "btn_Browse";
            this.btn_Browse.Size = new System.Drawing.Size(75, 28);
            this.btn_Browse.TabIndex = 6;
            this.btn_Browse.Text = "瀏覽";
            this.btn_Browse.UseVisualStyleBackColor = true;
            this.btn_Browse.Click += new System.EventHandler(this.btn_Browse_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // lb_UploadState
            // 
            this.lb_UploadState.AutoSize = true;
            this.lb_UploadState.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lb_UploadState.Location = new System.Drawing.Point(17, 211);
            this.lb_UploadState.Name = "lb_UploadState";
            this.lb_UploadState.Size = new System.Drawing.Size(0, 21);
            this.lb_UploadState.TabIndex = 7;
            // 
            // btn_IPConfirm
            // 
            this.btn_IPConfirm.Location = new System.Drawing.Point(324, 8);
            this.btn_IPConfirm.Name = "btn_IPConfirm";
            this.btn_IPConfirm.Size = new System.Drawing.Size(75, 28);
            this.btn_IPConfirm.TabIndex = 8;
            this.btn_IPConfirm.Text = "確認";
            this.btn_IPConfirm.UseVisualStyleBackColor = true;
            this.btn_IPConfirm.Click += new System.EventHandler(this.btn_IPConfirm_Click);
            // 
            // FTPUpload
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(417, 114);
            this.Controls.Add(this.btn_Upload);
            this.Controls.Add(this.btn_IPConfirm);
            this.Controls.Add(this.lb_UploadState);
            this.Controls.Add(this.btn_Browse);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.lb_Folder);
            this.Controls.Add(this.lb_IP);
            this.Controls.Add(this.tb_Folder);
            this.Controls.Add(this.tb_IP);
            this.Name = "FTPUpload";
            this.Text = "檔案上傳";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_IP;
        private System.Windows.Forms.TextBox tb_Folder;
        private System.Windows.Forms.Label lb_IP;
        private System.Windows.Forms.Label lb_Folder;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btn_Upload;
        private System.Windows.Forms.Button btn_Browse;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label lb_UploadState;
        private System.Windows.Forms.Button btn_IPConfirm;
    }
}