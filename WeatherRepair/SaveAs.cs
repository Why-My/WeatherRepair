using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WeatherRepair
{
    public partial class SaveAs : Form
    {
        public SaveAs()
        {
            InitializeComponent();
        }
      
        string ResourePath;
        string ExportPath;
        string SaveType;
        string OutPath;
        FileInfo[] WdataFile;
    
        public SaveAs(string ResourePath2)
        {
            InitializeComponent();
            ResourePath = ResourePath2;

        }
        private void BEXPORT_Click(object sender, EventArgs e)
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
                OUTtextBox2.Text = ExportPath;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SaveType = comboBox1.Text;
            OutPath = OUTtextBox2.Text;
            if (SaveType == "")
            {
                MessageBox.Show("请选择保存类型");

            }
            else if (OutPath == "")
            {
                MessageBox.Show("请选择输出路径");
              

            }
            else
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, SaveType, ExportPath);
                bar.ShowDialog();
            }
           

        }
        private void SaveAs_Load(object sender, EventArgs e)
        {

        }
    }
}
