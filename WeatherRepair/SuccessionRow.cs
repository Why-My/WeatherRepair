using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WeatherRepair
{
    public partial class SuccessionRow : Form
    {
        string ResourePath;
        string ExportPath;
        string OutPath;
        FileInfo[] WdataFile;
        public SuccessionRow(string ResourePath2)
        {
            InitializeComponent();
            ResourePath = ResourePath2;
        }
        protected bool IsInteger(string value)
        {
            string pattern = @"^[0-9]*[1-9][0-9]*$";
            return Regex.IsMatch(value, pattern);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string Days = textBox1.Text;
            bool DaysRight = IsInteger(Days);
            OutPath = OUTtextBox2.Text;
            if (!DaysRight)
            {
                MessageBox.Show("请输入正确格式参数");
            }
            else if (Days=="")
            {
                MessageBox.Show("请输入判断连续空缺天数");
            }
            else if(OutPath=="")
            {
                MessageBox.Show("请选择输出路径");
            }
            else
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, Days, ExportPath, "修补连续空缺日照时数");
                bar.ShowDialog();
            }
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
    }
}
