using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class ZRsunglight : Form
    {
        string ResourePath;
        FileInfo[] WdataFile;
        string  LatPath;
        public ZRsunglight()
        {
            InitializeComponent();
        }
        public ZRsunglight(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            bool a = false;
            bool b = false;
            string SDcols = textBox1.Text.Trim();
            if (SDcols == "")
            {
                MessageBox.Show("请输入列号");
            }
            else
            {
                if (IsWord(SDcols))
                {
                    a = true;
                }
                else
                {
                    MessageBox.Show("请输入正确格式列号");
                }
            }
            if (INPtextBox1.Text=="")
            {
                MessageBox.Show("请指定站点维度表");
            }
            else
            {
                b = true;
            }
            if (a&&b)
            {
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, SDcols, LatPath,6);
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

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "文件类型(*.xls,*.xlsx)|*.xls;*.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                LatPath = dialog.FileName;
                INPtextBox1.Text = LatPath;
            }
        }
    }
}
