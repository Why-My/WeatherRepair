namespace WeatherRepair
{
    partial class SaveAs
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SaveAs));
            this.label3 = new System.Windows.Forms.Label();
            this.BEXPORT = new System.Windows.Forms.Button();
            this.OUTtextBox2 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(53, 128);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 34;
            this.label3.Text = "选择输出文件夹";
            // 
            // BEXPORT
            // 
            this.BEXPORT.Location = new System.Drawing.Point(431, 153);
            this.BEXPORT.Name = "BEXPORT";
            this.BEXPORT.Size = new System.Drawing.Size(35, 21);
            this.BEXPORT.TabIndex = 33;
            this.BEXPORT.Text = "......";
            this.BEXPORT.UseVisualStyleBackColor = true;
            this.BEXPORT.Click += new System.EventHandler(this.BEXPORT_Click);
            // 
            // OUTtextBox2
            // 
            this.OUTtextBox2.Enabled = false;
            this.OUTtextBox2.Location = new System.Drawing.Point(55, 154);
            this.OUTtextBox2.Name = "OUTtextBox2";
            this.OUTtextBox2.Size = new System.Drawing.Size(367, 21);
            this.OUTtextBox2.TabIndex = 32;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(200, 253);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 30);
            this.button1.TabIndex = 31;
            this.button1.Text = "确定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "气象数据常规格式(*.dly)",
            "文本文件(制表符分隔)(*.txt)",
            "带格式文本文件(空格分隔)(*.prn)"});
            this.comboBox1.Location = new System.Drawing.Point(55, 92);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(367, 20);
            this.comboBox1.TabIndex = 30;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(53, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 29;
            this.label1.Text = "保存类型";
            // 
            // SaveAs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::WeatherRepair.Properties.Resources._34D4_E_61D2_WW_CY4_5FGT;
            this.ClientSize = new System.Drawing.Size(494, 301);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.BEXPORT);
            this.Controls.Add(this.OUTtextBox2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "SaveAs";
            this.Text = "另存为";
            this.Load += new System.EventHandler(this.SaveAs_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BEXPORT;
        private System.Windows.Forms.TextBox OUTtextBox2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
    }
}