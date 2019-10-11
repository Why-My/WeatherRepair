using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class Delete2 : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        public Delete2()
        {
            InitializeComponent();
        }
        public Delete2(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string SDcols = textBox3.Text.Trim();

            if (SDcols == "")
            {
                MessageBox.Show("请输入行号");
            }
            else
            {
                if (IsData(SDcols))
                {
                    DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                    WdataFile = Wdata.GetFiles();
                    ProgressBar bar = new ProgressBar(WdataFile, ResourePath, SDcols, 3);
                    bar.ShowDialog();
                }
                else
                {
                    MessageBox.Show("请输入正确格式行号");
                }
            }
        }
        private bool IsData(string value)
        {
            string col = value;
            string pattern2 = @"^[0-9]\d*$";
            bool IsData= Regex.IsMatch(col, pattern2);
            if (IsData)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
