using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WeatherRepair
{
    public partial class WeatherTool : Form
    {
        string ResourePath;
        string ExportPath;
        string Param;
        string OutPath;
        FileInfo[] WdataFile;
        public WeatherTool(string ResourePath2)
        {
            ResourePath = ResourePath2;
            InitializeComponent();
        }
        public WeatherTool()
        {
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("请选择目标文件夹");
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("请选择目标文件夹");
            }
            else
            {
                DeleteFile();
                Param = textBox2.Text.Trim();
                DirectoryInfo Wdata = new DirectoryInfo(ResourePath);
                WdataFile = Wdata.GetFiles();
                ProgressBar bar = new ProgressBar(WdataFile, ResourePath, ExportPath, Param,8);
                bar.ShowDialog();
            }
           
          
        }

        private void DeleteFile()
        {
          
            string[] DirName = { "Changedly", "inp", "out" };
            for (int i = 0; i < 3; i++)
            {
                string directoryPath = ExportPath + @"\" + DirName[i];
                //Directory.GetFiles(directoryPath).ToList().ForEach(File.Delete);
                if (!Directory.Exists(directoryPath))
                {
                    // Create the directory it does not exist.
                    Directory.CreateDirectory(directoryPath);
                }
                else
                {
                    Directory.Delete(directoryPath, true);
                    Directory.CreateDirectory(directoryPath);

                }
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ExportPath = dialog.SelectedPath;
                textBox1.Text = ExportPath;
            }
        }
    }
}
