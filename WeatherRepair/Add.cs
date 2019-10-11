using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class Add : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        string Row;
        string Column;
        public Add()
        {
            InitializeComponent();
        }
        public Add(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            bool a = false;
            bool b = false;
            bool c = false;
            Row = textBox1.Text.Trim();
            Column = textBox2.Text.Trim();
            string data2= textBox3.Text.Trim();
            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入行号");
            }
            else
            {
                if (IsData(Row))
                {
                    a = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的行号");
                }
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("请输入列号");
            }
            else
            {
                if (IsWord(Column) && Column.Length == 1)
                {
                    b = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的列号");
                }
            }
            if (textBox3.Text == "")
            {
                MessageBox.Show("请输入所添加的数据");
            }
            else
            {
                    c = true;
            }
            if (a && b && c)
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, Row, Column, data2,1);
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
