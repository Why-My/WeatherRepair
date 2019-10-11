using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class DleteFormulas : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        string MDcols;
        string GDcols;
        public DleteFormulas()
        {
            InitializeComponent();
        }
        public DleteFormulas(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            bool a = false;
            bool b = false;
            MDcols = textBox1.Text.Trim();
            GDcols = textBox2.Text.Trim();
            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入起始列号");
            }
            else
            {
                if (IsWord(MDcols) && MDcols.Length == 1)
                {
                    a = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的起始列号");
                }
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("请输入终止列号");
            }
            else
            {
                if (IsWord(GDcols) && GDcols.Length == 1)
                {
                    b = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的终止列号");
                }
            }
            if (a && b)
            {

                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, MDcols, GDcols, 5);
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
    }
}
