using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class AdjustColumn : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        string MDcols;
        string GDcols;
        public AdjustColumn()
        {
            InitializeComponent();
        }
        public AdjustColumn(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            bool a = false;
            bool b = false;
            bool c = false;
            MDcols = textBox1.Text.Trim();
            GDcols = textBox2.Text.Trim();
            string width= textBox3.Text.Trim();
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
            if (textBox3.Text=="")
            {
                MessageBox.Show("请输入列宽");
            }
            else
            {
                if (IsData(width))
                {
                    c = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式的列宽");
                }
            }
            if (a && b && c)
            {
                int width2 = Convert.ToInt32(width);
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, MDcols, GDcols, width2,"调整多列列宽");
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
