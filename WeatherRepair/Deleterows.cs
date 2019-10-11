using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class Deleterows : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        string MDcols;
        string GDcols;
        public Deleterows()
        {
            InitializeComponent();
        }
        public Deleterows(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void 确定_Click(object sender, EventArgs e)
        {
            bool a = false;
            bool b = false;
            MDcols = textBox1.Text.Trim();
            GDcols = textBox2.Text.Trim();
            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入起始行号");
            }
            else
            {
                if (IsData(MDcols))
                {
                    a = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的起始行号");
                }
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("请输入终止行号");
            }
            else
            {
                if (IsData(GDcols))
                {
                    b = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的终止行号");
                }
            }
            if (a && b)
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, MDcols, GDcols, 3);
                bar.ShowDialog();
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
