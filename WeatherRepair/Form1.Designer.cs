namespace WeatherRepair
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label2 = new System.Windows.Forms.Label();
            this.BINPUT = new System.Windows.Forms.Button();
            this.INPtextBox1 = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.文件ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.导入ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.另存为ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.气象数据修补ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修补连续空缺数据ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修补非连续空缺数据ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.调整ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.平移ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除单列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除多列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除单行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除多行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.调整列宽ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.调整单列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.调整多列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.添加数据ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修改公式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除公式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.清空单列公式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.清空多列公式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.添加公式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修改数据格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.计算逐日太阳辐射量ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.计算逐月气象时数ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.数据汇总ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.input汇总ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.output汇总ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.清除残留excel进程ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(79, 215);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 27;
            this.label2.Text = "选择目标资源";
            // 
            // BINPUT
            // 
            this.BINPUT.Location = new System.Drawing.Point(474, 250);
            this.BINPUT.Name = "BINPUT";
            this.BINPUT.Size = new System.Drawing.Size(28, 21);
            this.BINPUT.TabIndex = 26;
            this.BINPUT.Text = "......";
            this.BINPUT.UseVisualStyleBackColor = true;
            this.BINPUT.Click += new System.EventHandler(this.BINPUT_Click);
            // 
            // INPtextBox1
            // 
            this.INPtextBox1.Enabled = false;
            this.INPtextBox1.Location = new System.Drawing.Point(81, 250);
            this.INPtextBox1.Name = "INPtextBox1";
            this.INPtextBox1.Size = new System.Drawing.Size(387, 21);
            this.INPtextBox1.TabIndex = 25;
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件ToolStripMenuItem,
            this.气象数据修补ToolStripMenuItem,
            this.调整ToolStripMenuItem,
            this.toolStripMenuItem1,
            this.数据汇总ToolStripMenuItem,
            this.清除残留excel进程ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(584, 25);
            this.menuStrip1.TabIndex = 22;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 文件ToolStripMenuItem
            // 
            this.文件ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.导入ToolStripMenuItem,
            this.另存为ToolStripMenuItem});
            this.文件ToolStripMenuItem.Name = "文件ToolStripMenuItem";
            this.文件ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.文件ToolStripMenuItem.Text = "文件";
            // 
            // 导入ToolStripMenuItem
            // 
            this.导入ToolStripMenuItem.Name = "导入ToolStripMenuItem";
            this.导入ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.导入ToolStripMenuItem.Text = "导入";
            this.导入ToolStripMenuItem.Click += new System.EventHandler(this.导入ToolStripMenuItem_Click);
            // 
            // 另存为ToolStripMenuItem
            // 
            this.另存为ToolStripMenuItem.Name = "另存为ToolStripMenuItem";
            this.另存为ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.另存为ToolStripMenuItem.Text = "另存为";
            this.另存为ToolStripMenuItem.Click += new System.EventHandler(this.另存为ToolStripMenuItem_Click);
            // 
            // 气象数据修补ToolStripMenuItem
            // 
            this.气象数据修补ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.修补连续空缺数据ToolStripMenuItem,
            this.修补非连续空缺数据ToolStripMenuItem});
            this.气象数据修补ToolStripMenuItem.Name = "气象数据修补ToolStripMenuItem";
            this.气象数据修补ToolStripMenuItem.Size = new System.Drawing.Size(92, 21);
            this.气象数据修补ToolStripMenuItem.Text = "气象数据修补";
            // 
            // 修补连续空缺数据ToolStripMenuItem
            // 
            this.修补连续空缺数据ToolStripMenuItem.Name = "修补连续空缺数据ToolStripMenuItem";
            this.修补连续空缺数据ToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.修补连续空缺数据ToolStripMenuItem.Text = "修补连续空缺数据";
            this.修补连续空缺数据ToolStripMenuItem.Click += new System.EventHandler(this.修补连续空缺数据ToolStripMenuItem_Click);
            // 
            // 修补非连续空缺数据ToolStripMenuItem
            // 
            this.修补非连续空缺数据ToolStripMenuItem.Name = "修补非连续空缺数据ToolStripMenuItem";
            this.修补非连续空缺数据ToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.修补非连续空缺数据ToolStripMenuItem.Text = "修补非连续空缺数据";
            this.修补非连续空缺数据ToolStripMenuItem.Click += new System.EventHandler(this.修补非连续空缺数据ToolStripMenuItem_Click);
            // 
            // 调整ToolStripMenuItem
            // 
            this.调整ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.平移ToolStripMenuItem,
            this.删除ToolStripMenuItem,
            this.调整列宽ToolStripMenuItem,
            this.添加数据ToolStripMenuItem,
            this.修改公式ToolStripMenuItem,
            this.修改数据格式ToolStripMenuItem});
            this.调整ToolStripMenuItem.Name = "调整ToolStripMenuItem";
            this.调整ToolStripMenuItem.Size = new System.Drawing.Size(116, 21);
            this.调整ToolStripMenuItem.Text = "气象数据格式调整";
            // 
            // 平移ToolStripMenuItem
            // 
            this.平移ToolStripMenuItem.Name = "平移ToolStripMenuItem";
            this.平移ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.平移ToolStripMenuItem.Text = "平移列";
            this.平移ToolStripMenuItem.Click += new System.EventHandler(this.平移ToolStripMenuItem_Click);
            // 
            // 删除ToolStripMenuItem
            // 
            this.删除ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.删除列ToolStripMenuItem,
            this.删除行ToolStripMenuItem});
            this.删除ToolStripMenuItem.Name = "删除ToolStripMenuItem";
            this.删除ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除ToolStripMenuItem.Text = "删除数据";
            this.删除ToolStripMenuItem.Click += new System.EventHandler(this.删除ToolStripMenuItem_Click);
            // 
            // 删除列ToolStripMenuItem
            // 
            this.删除列ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.删除单列ToolStripMenuItem,
            this.删除多列ToolStripMenuItem});
            this.删除列ToolStripMenuItem.Name = "删除列ToolStripMenuItem";
            this.删除列ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.删除列ToolStripMenuItem.Text = "删除列";
            this.删除列ToolStripMenuItem.Click += new System.EventHandler(this.删除列ToolStripMenuItem_Click);
            // 
            // 删除单列ToolStripMenuItem
            // 
            this.删除单列ToolStripMenuItem.Name = "删除单列ToolStripMenuItem";
            this.删除单列ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除单列ToolStripMenuItem.Text = "删除单列";
            this.删除单列ToolStripMenuItem.Click += new System.EventHandler(this.删除单列ToolStripMenuItem_Click);
            // 
            // 删除多列ToolStripMenuItem
            // 
            this.删除多列ToolStripMenuItem.Name = "删除多列ToolStripMenuItem";
            this.删除多列ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除多列ToolStripMenuItem.Text = "删除多列";
            this.删除多列ToolStripMenuItem.Click += new System.EventHandler(this.删除多列ToolStripMenuItem_Click);
            // 
            // 删除行ToolStripMenuItem
            // 
            this.删除行ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.删除单行ToolStripMenuItem,
            this.删除多行ToolStripMenuItem});
            this.删除行ToolStripMenuItem.Name = "删除行ToolStripMenuItem";
            this.删除行ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.删除行ToolStripMenuItem.Text = "删除行";
            this.删除行ToolStripMenuItem.Click += new System.EventHandler(this.删除行ToolStripMenuItem_Click);
            // 
            // 删除单行ToolStripMenuItem
            // 
            this.删除单行ToolStripMenuItem.Name = "删除单行ToolStripMenuItem";
            this.删除单行ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除单行ToolStripMenuItem.Text = "删除单行";
            this.删除单行ToolStripMenuItem.Click += new System.EventHandler(this.删除单行ToolStripMenuItem_Click);
            // 
            // 删除多行ToolStripMenuItem
            // 
            this.删除多行ToolStripMenuItem.Name = "删除多行ToolStripMenuItem";
            this.删除多行ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除多行ToolStripMenuItem.Text = "删除多行";
            this.删除多行ToolStripMenuItem.Click += new System.EventHandler(this.删除多行ToolStripMenuItem_Click);
            // 
            // 调整列宽ToolStripMenuItem
            // 
            this.调整列宽ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.调整单列ToolStripMenuItem,
            this.调整多列ToolStripMenuItem});
            this.调整列宽ToolStripMenuItem.Name = "调整列宽ToolStripMenuItem";
            this.调整列宽ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.调整列宽ToolStripMenuItem.Text = "调整列宽";
            this.调整列宽ToolStripMenuItem.Click += new System.EventHandler(this.调整列宽ToolStripMenuItem_Click);
            // 
            // 调整单列ToolStripMenuItem
            // 
            this.调整单列ToolStripMenuItem.Name = "调整单列ToolStripMenuItem";
            this.调整单列ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.调整单列ToolStripMenuItem.Text = "调整单列";
            this.调整单列ToolStripMenuItem.Click += new System.EventHandler(this.调整单列ToolStripMenuItem_Click);
            // 
            // 调整多列ToolStripMenuItem
            // 
            this.调整多列ToolStripMenuItem.Name = "调整多列ToolStripMenuItem";
            this.调整多列ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.调整多列ToolStripMenuItem.Text = "调整多列";
            this.调整多列ToolStripMenuItem.Click += new System.EventHandler(this.调整多列ToolStripMenuItem_Click);
            // 
            // 添加数据ToolStripMenuItem
            // 
            this.添加数据ToolStripMenuItem.Name = "添加数据ToolStripMenuItem";
            this.添加数据ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.添加数据ToolStripMenuItem.Text = "添加数据";
            this.添加数据ToolStripMenuItem.Click += new System.EventHandler(this.添加数据ToolStripMenuItem_Click);
            // 
            // 修改公式ToolStripMenuItem
            // 
            this.修改公式ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.删除公式ToolStripMenuItem,
            this.添加公式ToolStripMenuItem});
            this.修改公式ToolStripMenuItem.Name = "修改公式ToolStripMenuItem";
            this.修改公式ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.修改公式ToolStripMenuItem.Text = "修改公式";
            this.修改公式ToolStripMenuItem.Click += new System.EventHandler(this.修改公式ToolStripMenuItem_Click);
            // 
            // 删除公式ToolStripMenuItem
            // 
            this.删除公式ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.清空单列公式ToolStripMenuItem,
            this.清空多列公式ToolStripMenuItem});
            this.删除公式ToolStripMenuItem.Name = "删除公式ToolStripMenuItem";
            this.删除公式ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除公式ToolStripMenuItem.Text = "清空公式";
            this.删除公式ToolStripMenuItem.Click += new System.EventHandler(this.删除公式ToolStripMenuItem_Click);
            // 
            // 清空单列公式ToolStripMenuItem
            // 
            this.清空单列公式ToolStripMenuItem.Name = "清空单列公式ToolStripMenuItem";
            this.清空单列公式ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.清空单列公式ToolStripMenuItem.Text = "清空单列公式";
            this.清空单列公式ToolStripMenuItem.Click += new System.EventHandler(this.清空单列公式ToolStripMenuItem_Click);
            // 
            // 清空多列公式ToolStripMenuItem
            // 
            this.清空多列公式ToolStripMenuItem.Name = "清空多列公式ToolStripMenuItem";
            this.清空多列公式ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.清空多列公式ToolStripMenuItem.Text = "清空多列公式";
            this.清空多列公式ToolStripMenuItem.Click += new System.EventHandler(this.清空多列公式ToolStripMenuItem_Click);
            // 
            // 添加公式ToolStripMenuItem
            // 
            this.添加公式ToolStripMenuItem.Name = "添加公式ToolStripMenuItem";
            this.添加公式ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.添加公式ToolStripMenuItem.Text = "添加公式";
            this.添加公式ToolStripMenuItem.Click += new System.EventHandler(this.添加公式ToolStripMenuItem_Click);
            // 
            // 修改数据格式ToolStripMenuItem
            // 
            this.修改数据格式ToolStripMenuItem.Name = "修改数据格式ToolStripMenuItem";
            this.修改数据格式ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.修改数据格式ToolStripMenuItem.Text = "修改格式";
            this.修改数据格式ToolStripMenuItem.Click += new System.EventHandler(this.修改数据格式ToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.计算逐日太阳辐射量ToolStripMenuItem,
            this.计算逐月气象时数ToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(92, 21);
            this.toolStripMenuItem1.Text = "气象数据处理";
            // 
            // 计算逐日太阳辐射量ToolStripMenuItem
            // 
            this.计算逐日太阳辐射量ToolStripMenuItem.Name = "计算逐日太阳辐射量ToolStripMenuItem";
            this.计算逐日太阳辐射量ToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.计算逐日太阳辐射量ToolStripMenuItem.Text = "计算逐日太阳辐射量";
            this.计算逐日太阳辐射量ToolStripMenuItem.Click += new System.EventHandler(this.计算逐日太阳辐射量ToolStripMenuItem_Click);
            // 
            // 计算逐月气象时数ToolStripMenuItem
            // 
            this.计算逐月气象时数ToolStripMenuItem.Name = "计算逐月气象时数ToolStripMenuItem";
            this.计算逐月气象时数ToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.计算逐月气象时数ToolStripMenuItem.Text = "WeatherTool";
            this.计算逐月气象时数ToolStripMenuItem.Click += new System.EventHandler(this.计算逐月气象时数ToolStripMenuItem_Click);
            // 
            // 数据汇总ToolStripMenuItem
            // 
            this.数据汇总ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.input汇总ToolStripMenuItem,
            this.output汇总ToolStripMenuItem});
            this.数据汇总ToolStripMenuItem.Name = "数据汇总ToolStripMenuItem";
            this.数据汇总ToolStripMenuItem.Size = new System.Drawing.Size(68, 21);
            this.数据汇总ToolStripMenuItem.Text = "数据汇总";
            // 
            // input汇总ToolStripMenuItem
            // 
            this.input汇总ToolStripMenuItem.Name = "input汇总ToolStripMenuItem";
            this.input汇总ToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.input汇总ToolStripMenuItem.Text = "input汇总";
            this.input汇总ToolStripMenuItem.Click += new System.EventHandler(this.input汇总ToolStripMenuItem_Click);
            // 
            // output汇总ToolStripMenuItem
            // 
            this.output汇总ToolStripMenuItem.Name = "output汇总ToolStripMenuItem";
            this.output汇总ToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.output汇总ToolStripMenuItem.Text = "output汇总";
            this.output汇总ToolStripMenuItem.Click += new System.EventHandler(this.output汇总ToolStripMenuItem_Click);
            // 
            // 清除残留excel进程ToolStripMenuItem
            // 
            this.清除残留excel进程ToolStripMenuItem.Name = "清除残留excel进程ToolStripMenuItem";
            this.清除残留excel进程ToolStripMenuItem.Size = new System.Drawing.Size(121, 21);
            this.清除残留excel进程ToolStripMenuItem.Text = "清除残留excel进程";
            this.清除残留excel进程ToolStripMenuItem.Click += new System.EventHandler(this.清除残留excel进程ToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::WeatherRepair.Properties.Resources.图层_5;
            this.ClientSize = new System.Drawing.Size(584, 376);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BINPUT);
            this.Controls.Add(this.INPtextBox1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "气象数据修补处理系统";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button BINPUT;
        private System.Windows.Forms.TextBox INPtextBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 文件ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 导入ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 另存为ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 气象数据修补ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修补连续空缺数据ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修补非连续空缺数据ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 调整ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 平移ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem 计算逐日太阳辐射量ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 计算逐月气象时数ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 数据汇总ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem input汇总ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem output汇总ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 调整列宽ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除列ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除行ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 添加数据ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修改公式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修改数据格式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 调整单列ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 调整多列ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除单列ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除多列ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除单行ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除多行ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除公式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 添加公式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 清空单列公式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 清空多列公式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 清除残留excel进程ToolStripMenuItem;
    }
}

