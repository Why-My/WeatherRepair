using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class Move : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        public Move()
        {
            InitializeComponent();
        }
        public Move(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            IsNull();
            string Mcols = textBox1.Text.Trim();
            string Gcols = textBox2.Text.Trim();
      
            if (!IsWord(Mcols))
            {
                MessageBox.Show("请输入正确的移动列号");
            }
            else if (!IsWord(Gcols))
            {
                MessageBox.Show("请输入正确的插入列号");
            }
            else
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, Mcols, Gcols,1);
                bar.ShowDialog();
            }   
         
        }
        protected bool IsWord(string value)
        {
            string col = value.Trim();
            string pattern = @"^[A-Za-z]{1}$";
            bool Isword =Regex.IsMatch(col, pattern);
            if (col.Length == 1 && Isword)
            {
                return true;
               
            }
            else
            {
                return false;
               

            }
         
            
        }
        private void IsNull()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入移动列号");
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("请输入插入列号");
            }
          

        }
    }
}
