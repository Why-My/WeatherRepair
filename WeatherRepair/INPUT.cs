using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WeatherRepair
{
    public partial class INPUT : Form
    {
        string ResourePath;
        string ExportPath;
        string OutPath;
        FileInfo[] WdataFile;
        public INPUT()
        {
            InitializeComponent();
        }
        public INPUT(string ResourePath2 )
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            OutPath = textBox1.Text;
            if (OutPath == "")
            {
                MessageBox.Show("请输入正确的输出路径");
            }
            else
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, ExportPath, 0.2);
                bar.ShowDialog();
            }
        }
        private void INPUT_Load(object sender, EventArgs e)
        {

        }
       

        private void button1_Click_1(object sender, EventArgs e)
        {

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (ExportPath != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog.SelectedPath = ExportPath;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ExportPath = dialog.SelectedPath;
                textBox1.Text = ExportPath;
            }
        }
    }
}
