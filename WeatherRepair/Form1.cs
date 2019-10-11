using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
namespace WeatherRepair
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        struct FilePath
        {
            public string ResourePath;//文件夹路径
            public string FilesPath;//文件路径
            public string ExportPath;//输出路径
            public string FilePath2;//外部文件路径BAOJI2.XLS
            public string PrnPath;//prn路径
            public string InputPath;//input路径
            public string OutPath;//output路径
        }
        FilePath filePath1;
        int Finially;
        string Debug = System.Windows.Forms.Application.StartupPath;
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hwnd, out int processid);
        private delegate void setTextValueCallBack(string wFileName);
        private delegate void setTextValueCallBack2();
        private setTextValueCallBack setCallBack;
        private setTextValueCallBack2 setCallBack2;
        private setTextValueCallBack2 setCallBack3;
        FileInfo[] WdataFile;
        //WriteExcel WriteExcel;
        Thread thread1;
        private void 导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (filePath1.ResourePath != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog.SelectedPath = filePath1.ResourePath;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath1.ResourePath = dialog.SelectedPath;
                INPtextBox1.Text = filePath1.ResourePath;
            }
        }

        private void 另存为ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    SaveAs save = new SaveAs(filePath1.ResourePath);
                    save.ShowDialog();
                    break;
            }
        }
        private int PathRight()
        {
            if (INPtextBox1.Text == "")
            {
                return 0;
            }

            return 2;
        }//判断输入输出路径是否存在
        
        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void BINPUT_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath1.ResourePath = dialog.SelectedPath;
                INPtextBox1.Text = filePath1.ResourePath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
     
        private void 修补连续空缺数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                   SuccessionRow SuccessionRow = new SuccessionRow(filePath1.ResourePath);
                    SuccessionRow.ShowDialog();
                    break;
            }
        }

        private void 修补非连续空缺数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    SingleRow SingleRow = new SingleRow(filePath1.ResourePath);
                    SingleRow.ShowDialog();
                    break;
            }
        }

        private void 平移ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Move Move = new Move(filePath1.ResourePath);
                    Move.ShowDialog();
                    break;
            }
        }


        private void 添加列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Add Add = new Add(filePath1.ResourePath);
                    Add.ShowDialog();
                    break;
            }
        }

        private void 删除列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void 删除行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           

        }

        private void 调整列宽ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }

        private void 添加数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Add Add = new Add(filePath1.ResourePath);
                    Add.ShowDialog();
                    break;
            }

        }

        private void 调整单列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    ColumnWidth ColumnWidth = new ColumnWidth(filePath1.ResourePath);
                    ColumnWidth.ShowDialog();
                    break;
            }
        }

        private void 调整多列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    AdjustColumn ColumnWidth = new AdjustColumn(filePath1.ResourePath);
                    ColumnWidth.ShowDialog();
                    break;
            }
        }

        private void 删除单行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Delete2 Delete2 = new Delete2(filePath1.ResourePath);
                    Delete2.ShowDialog();
                    break;
            }

        }

        private void 删除单列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Delete Delete = new Delete(filePath1.ResourePath);
                    Delete.ShowDialog();
                    break;
            }
        }

        private void 删除多列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Deletecolunms Deletes = new Deletecolunms(filePath1.ResourePath);
                    Deletes.ShowDialog();
                    break;
            }
        }

        private void 删除多行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Deleterows Deleterows = new Deleterows(filePath1.ResourePath);
                    Deleterows.ShowDialog();
                    break;
            }

        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 修改数据格式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    NumberType NumberType = new NumberType(filePath1.ResourePath);
                    NumberType.ShowDialog();
                    break;
            }
        }

        private void 删除公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void 添加公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    Formula Formula = new Formula(filePath1.ResourePath);
                    Formula.ShowDialog();
                    break;
            }
        }

        private void 清空单列公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    DeleteFormula DeleteFormula = new DeleteFormula(filePath1.ResourePath);
                    DeleteFormula.ShowDialog();
                    break;
            }
        }

        private void 清空多列公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    DleteFormulas DeleteFormulas = new DleteFormulas(filePath1.ResourePath);
                    DeleteFormulas.ShowDialog();
                    break;
            }
        }

        private void 修改公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 计算逐日太阳辐射量ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                   ZRsunglight ZRsunglight = new ZRsunglight(filePath1.ResourePath);
                    ZRsunglight.ShowDialog();
                    break;
            }
        }

        private void 计算逐月气象时数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    WeatherTool WeatherTool = new WeatherTool(filePath1.ResourePath);
                    WeatherTool.ShowDialog();
                    break;
            }
        }

        private void input汇总ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    INPUT INPUT = new INPUT (filePath1.ResourePath);
                    INPUT.ShowDialog();
                    break;
            }
        }

        private void output汇总ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int pathright = PathRight();
            switch (pathright)
            {
                case 0:
                    MessageBox.Show("请确认输入路径");
                    break;
                default:
                    OUTPUT OUTPUT = new OUTPUT(filePath1.ResourePath);
                    OUTPUT.ShowDialog();
                    break;
            }
        }

        private void 清除残留excel进程ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("是否清除所有excel进程？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.OK)
            {
                DialogResult dr2 = MessageBox.Show("请在清除前将正在使用的excel保存，以免数据丢失", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr == DialogResult.OK)
                {
                    KillExcel();
                    MessageBox.Show("清除成功");
                }


            }
           
        }

        private void KillExcel()
        {
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
    }
}
