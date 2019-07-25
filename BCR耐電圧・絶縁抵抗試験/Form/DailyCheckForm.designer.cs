namespace BCR耐電圧_絶縁抵抗試験
{
    partial class DailyCheckForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DailyCheckForm));
            this.buttonReturn = new System.Windows.Forms.Button();
            this.labelDecision = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.labelErrorMess = new System.Windows.Forms.Label();
            this.buttonStart = new System.Windows.Forms.Button();
            this.labelBcrCh1Ch4Check = new System.Windows.Forms.Label();
            this.labelBcrCh23Ch5Check = new System.Windows.Forms.Label();
            this.labelMessage = new System.Windows.Forms.Label();
            this.labelDanger = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.labelAurCh1Ch2Check = new System.Windows.Forms.Label();
            this.timerLbMessage = new System.Windows.Forms.Timer(this.components);
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.timerCount = new System.Windows.Forms.Timer(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonReturn
            // 
            this.buttonReturn.Location = new System.Drawing.Point(12, 133);
            this.buttonReturn.Name = "buttonReturn";
            this.buttonReturn.Size = new System.Drawing.Size(108, 29);
            this.buttonReturn.TabIndex = 0;
            this.buttonReturn.Text = "メインに戻る";
            this.buttonReturn.UseVisualStyleBackColor = true;
            this.buttonReturn.Click += new System.EventHandler(this.buttonReturn_Click);
            // 
            // labelDecision
            // 
            this.labelDecision.BackColor = System.Drawing.Color.Transparent;
            this.labelDecision.Font = new System.Drawing.Font("メイリオ", 99.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelDecision.ForeColor = System.Drawing.Color.LightSeaGreen;
            this.labelDecision.Location = new System.Drawing.Point(23, 0);
            this.labelDecision.Name = "labelDecision";
            this.labelDecision.Size = new System.Drawing.Size(420, 159);
            this.labelDecision.TabIndex = 1;
            this.labelDecision.Text = "PASS";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelErrorMess);
            this.groupBox1.Controls.Add(this.labelDecision);
            this.groupBox1.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox1.Location = new System.Drawing.Point(363, 117);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(481, 213);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "判定";
            // 
            // labelErrorMess
            // 
            this.labelErrorMess.AutoSize = true;
            this.labelErrorMess.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelErrorMess.ForeColor = System.Drawing.Color.Maroon;
            this.labelErrorMess.Location = new System.Drawing.Point(52, 168);
            this.labelErrorMess.Name = "labelErrorMess";
            this.labelErrorMess.Size = new System.Drawing.Size(148, 28);
            this.labelErrorMess.TabIndex = 7;
            this.labelErrorMess.Text = "labelErrorMess";
            // 
            // buttonStart
            // 
            this.buttonStart.Font = new System.Drawing.Font("メイリオ", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.buttonStart.Location = new System.Drawing.Point(139, 121);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(149, 47);
            this.buttonStart.TabIndex = 3;
            this.buttonStart.Text = "開始";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // labelBcrCh1Ch4Check
            // 
            this.labelBcrCh1Ch4Check.AutoSize = true;
            this.labelBcrCh1Ch4Check.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelBcrCh1Ch4Check.ForeColor = System.Drawing.Color.Black;
            this.labelBcrCh1Ch4Check.Location = new System.Drawing.Point(16, 39);
            this.labelBcrCh1Ch4Check.Name = "labelBcrCh1Ch4Check";
            this.labelBcrCh1Ch4Check.Size = new System.Drawing.Size(259, 28);
            this.labelBcrCh1Ch4Check.TabIndex = 5;
            this.labelBcrCh1Ch4Check.Text = "①BCR側 CH1-CH4導通ﾁｪｯｸ";
            // 
            // labelBcrCh23Ch5Check
            // 
            this.labelBcrCh23Ch5Check.AutoSize = true;
            this.labelBcrCh23Ch5Check.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelBcrCh23Ch5Check.ForeColor = System.Drawing.Color.Black;
            this.labelBcrCh23Ch5Check.Location = new System.Drawing.Point(16, 79);
            this.labelBcrCh23Ch5Check.Name = "labelBcrCh23Ch5Check";
            this.labelBcrCh23Ch5Check.Size = new System.Drawing.Size(278, 28);
            this.labelBcrCh23Ch5Check.TabIndex = 6;
            this.labelBcrCh23Ch5Check.Text = "②BCR側 CH2,3-CH5導通ﾁｪｯｸ";
            // 
            // labelMessage
            // 
            this.labelMessage.BackColor = System.Drawing.Color.LightBlue;
            this.labelMessage.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelMessage.ForeColor = System.Drawing.Color.Black;
            this.labelMessage.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.labelMessage.Location = new System.Drawing.Point(13, 32);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(876, 48);
            this.labelMessage.TabIndex = 7;
            this.labelMessage.Text = "点検用の製品をセットして開始ボタンを押してください";
            this.labelMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labelDanger
            // 
            this.labelDanger.BackColor = System.Drawing.Color.MistyRose;
            this.labelDanger.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelDanger.Font = new System.Drawing.Font("MS UI Gothic", 65.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelDanger.ForeColor = System.Drawing.Color.White;
            this.labelDanger.Location = new System.Drawing.Point(12, 9);
            this.labelDanger.Name = "labelDanger";
            this.labelDanger.Size = new System.Drawing.Size(832, 104);
            this.labelDanger.TabIndex = 8;
            this.labelDanger.Text = "高 電 圧 注 意 ！";
            this.labelDanger.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox5
            // 
            this.groupBox5.BackColor = System.Drawing.Color.Transparent;
            this.groupBox5.Controls.Add(this.labelMessage);
            this.groupBox5.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox5.Location = new System.Drawing.Point(12, 332);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(895, 91);
            this.groupBox5.TabIndex = 11;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "操作指示";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.labelAurCh1Ch2Check);
            this.groupBox2.Controls.Add(this.labelBcrCh1Ch4Check);
            this.groupBox2.Controls.Add(this.labelBcrCh23Ch5Check);
            this.groupBox2.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox2.Location = new System.Drawing.Point(12, 174);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(345, 152);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "検査項目";
            // 
            // labelAurCh1Ch2Check
            // 
            this.labelAurCh1Ch2Check.AutoSize = true;
            this.labelAurCh1Ch2Check.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.labelAurCh1Ch2Check.ForeColor = System.Drawing.Color.Black;
            this.labelAurCh1Ch2Check.Location = new System.Drawing.Point(16, 116);
            this.labelAurCh1Ch2Check.Name = "labelAurCh1Ch2Check";
            this.labelAurCh1Ch2Check.Size = new System.Drawing.Size(260, 28);
            this.labelAurCh1Ch2Check.TabIndex = 7;
            this.labelAurCh1Ch2Check.Text = "③AUR側 CH1-CH2導通ﾁｪｯｸ";
            // 
            // timerLbMessage
            // 
            this.timerLbMessage.Tick += new System.EventHandler(this.timerLbMessage_Tick);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 441);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(345, 239);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
            // 
            // timerCount
            // 
            this.timerCount.Tick += new System.EventHandler(this.timerCount_Tick);
            // 
            // DailyCheckForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(919, 690);
            this.ControlBox = false;
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.labelDanger);
            this.Controls.Add(this.buttonStart);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonReturn);
            this.Name = "DailyCheckForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "日常点検";
            this.Load += new System.EventHandler(this.DailyCheckForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonReturn;
        private System.Windows.Forms.Label labelDecision;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.Label labelBcrCh1Ch4Check;
        private System.Windows.Forms.Label labelBcrCh23Ch5Check;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Label labelDanger;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Timer timerLbMessage;
        private System.Windows.Forms.Label labelErrorMess;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelAurCh1Ch2Check;
        private System.Windows.Forms.Timer timerCount;
    }
}