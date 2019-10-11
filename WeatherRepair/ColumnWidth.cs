using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class ColumnWidth : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        string SDcols;
        public ColumnWidth()
        {
            InitializeComponent();
        }
        public ColumnWidth(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            bool a = false;
            bool b = false;
            SDcols = textBox1.Text.Trim();
           string width =textBox2.Text.Trim()  ;
            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入列号");
            }
            else
            {
                if (IsWord(SDcols)&& SDcols.Length==1)
                {
                    a = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式列号");
                }
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("请输入列宽");
            }
            else
            {
                if (IsData(width))
                {
                    b= true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式列宽");
                }
            }

            if (a&&b)
            {
                int width2 = Convert.ToInt32(width);
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, SDcols, width2,"调整单列列宽");
                bar.ShowDialog();
            }
        }
        protected bool IsWord(string value)
        {
            string col = value;
            string pattern = @"^[A-Za-z]{1}$";
            bool Isword = Regex.IsMatch(col, pattern);
            if (col.Length == 1 && Isword)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool IsData(string value)
        {
            string col = value;
            string pattern2 = @"^[0-9]\d*$";
            bool Isword = Regex.IsMatch(col, pattern2);
            if (Isword)
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
