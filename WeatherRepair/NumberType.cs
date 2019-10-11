using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class NumberType : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        public NumberType()
        {
            InitializeComponent();
        }
        public NumberType(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string SDcols = textBox1.Text.Trim();
            if (SDcols == "")
            {
                MessageBox.Show("请输入列号");
            }
            else
            {
                if (IsWord(SDcols))
                {
                    DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                    WdataFile = Wdata.GetFiles();
                    ProgressBar bar = new ProgressBar(WdataFile, ResourePath, SDcols, 5);
                    bar.ShowDialog();
                }
                else
                {
                    MessageBox.Show("请输入正确格式列号");
                }
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
    }
}
