using System;
using System.IO;
using System.Windows.Forms;

namespace WeatherRepair
{
    public partial class SingleRow : Form
    {
        public SingleRow()
        {
            InitializeComponent();
        }

        public SingleRow(string ResourePath2)
        {
            InitializeComponent();
            ResourePath = ResourePath2;
        }
        string ResourePath;
        string ExportPath;
        string OutPath;
        FileInfo[] WdataFile;
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
            OutPath = OUTtextBox2.Text;
            if (OutPath == "")
            {
                MessageBox.Show("请输入正确的输出路径");
            }
            else
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, ExportPath, 0.1);
                bar.ShowDialog();
            }
        }
    }
}
