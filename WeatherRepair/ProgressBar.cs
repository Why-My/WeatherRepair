using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace WeatherRepair
{
    public partial class ProgressBar : Form
    {
        public ProgressBar()
        {
            InitializeComponent();
        }
        FileInfo[] WdataFile;
        public Thread thread1;
        string ResourePath;
        string SaveType;
        string Days;
        string ExportPath;
        string Mcols;
        string Gcols;
        string Scols;
        string Row;
        string Column;
        string Data;
        string param;
        int width;
        string LatPath;
        bool Stop=true;
        string Formula;
        int count;
        int Span=1;
        string HZPath;
        Excel.Workbook xlsWorkBook3;//inp
        Excel.Workbook xlsWorkBook4;//out
        Excel.Worksheet xlsWorkSheet3;
        Excel.Worksheet xlsWorkSheet4;
        ExcelProcess ExcelProcess2 ;
        private delegate void setTextValueCallBack(string wFileName);//对显示处理文件名方法的的委托
        private delegate void setExcelWay(string FilesPath);//对Excel操作委托
        private delegate void setTextValueCallBack2();//对进度条增长的委托
        private setTextValueCallBack setCallBack;
        private setTextValueCallBack2 setCallBack2;
        private setExcelWay ExcelWay;
        public ProgressBar(FileInfo[] WdataFile2,string ResourePath2,string SaveType2,string ExportPath2)
        {
            SaveType = SaveType2;
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            ExcelWay = new setExcelWay(SaveAs2);
            ExportPath = ExportPath2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;
        }//另存重载
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string Days2, string ExportPath2,string a)
        {
            Days = Days2;
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            ExcelWay= new setExcelWay(SuccessionRow);
            ExportPath = ExportPath2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;


        }//连续空缺天数重载
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string ExportPath2,double a)
        {
            InitializeComponent();
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            switch (a)
            { case 0.1:
                    ExcelWay = new setExcelWay(SingleRow);
                    break;
                case 0.2:
                    ExcelWay = new setExcelWay(INPUT);
                    break;
                case 0.3:
                    ExcelWay = new setExcelWay(OUTPUT);
                    break;
                default:
                    break;
            }
            ExportPath = ExportPath2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            ExcelProcess2 = new ExcelProcess();
            timer1.Enabled = true;
        }//非连续空缺天数重载
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string Mcols2, string Gcols2,int a)
        {
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            switch (a)
            {
                case 1:
                    ExcelWay = new setExcelWay(MoveCols);
                    break;
                case 2:
                    ExcelWay = new setExcelWay(DeleteCols);
                    break;
                case 3:
                    ExcelWay = new setExcelWay(DeleteRows);
                    break;
                case 4:
                    ExcelWay = new setExcelWay(AddFormula);
                    break;
                case 5:
                    ExcelWay = new setExcelWay(DeleteFormulas);
                    break;
                case 6:
                    ExcelWay = new setExcelWay(ZSSunlight);
                    break;
                case 8:
                    ExcelWay = new setExcelWay(WeatherTool);
                    break;
                default:
                    break;
            }
            LatPath = Gcols2;
            Column = Mcols2;
            Formula = Gcols2;
            width = a;
            Mcols = Mcols2;
            Gcols = Gcols2;
            param = Gcols2;
            ResourePath = ResourePath2;
            ExportPath = Mcols2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;
        }//移动列的重载/删除多列/删除多行
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string Mcols2, string Gcols2, int width2,string a)
        {
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            ExcelWay = new setExcelWay(AdjustColumns);
            width = width2;
            Mcols = Mcols2;
            Gcols = Gcols2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;
        }//调整连续列列宽
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string Scols2, int a)
        {
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            switch (a)
            {
                case 2:
                    ExcelWay = new setExcelWay(DeleteCol);
                    break;
                case 3:
                    ExcelWay = new setExcelWay(DeleteRow);
                    break;
                case 4:
                    ExcelWay = new setExcelWay(DeleteFormula);
                    break;
                case 5:
                    ExcelWay = new setExcelWay(ChangeType);
                    break;
                default:
                    break;
            }
            width = a;
            Scols = Scols2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;

        }//删除单列//删除单行
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string Scols2, int width2,string a)
        {
            setCallBack = new setTextValueCallBack(SetValue);
            ExcelWay = new setExcelWay(AdjustColumn);
            width = width2;
            Scols = Scols2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;
        }//调整单列列宽
        public ProgressBar(FileInfo[] WdataFile2, string ResourePath2, string Row2,string Column2,string Data2, int c)
        {
            setCallBack = new setTextValueCallBack(SetValue);
            setCallBack2 = new setTextValueCallBack2(ChangeProgress);
            ExcelWay = new setExcelWay(AddData);
            Row = Row2;
            Column = Column2;
            Data = Data2;
            ResourePath = ResourePath2;
            WdataFile = WdataFile2;
            InitializeComponent();
            timer1.Enabled = true;
        }//添加数据
        private void SaveAs2(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.SaveAs2(SaveType);
            count++;
        }//另存方法
        private void SuccessionRow(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.SuccessionRow(Days);
            count++;
        }//连续空缺方法
        private void SingleRow(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.SingleRow();
            count++;
        }//非连续空缺方法
        private void MoveCols(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.MoveClos(Mcols, Gcols);
            count++;

        }//移动列的方法
        private void DeleteCols(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.DeleteCols(Mcols, Gcols);
            count++;
        }//删除列的方法
        private void ZSSunlight(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.ZSSunlight(LatPath, Mcols);
            count++;
        }//删除列的方法
        private void DeleteFormulas(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.DeleteFormulas (Mcols, Gcols);
            count++;
        }//删除公式的方法
        private void DeleteCol(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.DeleteCol(Scols);
            count++;
        }//删除列的方法
        private void ChangeType(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.ChangeType(Scols);
            count++;
        }//改变数据格式的方法
        private void DeleteFormula(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.DeleteFormula(Scols);
            count++;
        }//删除公式的方法
        private void DeleteRows(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.DeleteRows(Mcols, Gcols);
            count++;
        }//删除连续多行的方法
        private void DeleteRow(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.DeleteRow(Scols);
            count++;
        }//删除单行的方法
        private void AdjustColumn(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.Adjust(Scols, width);
            count++;
        }//调整列宽的方法
        private void AdjustColumns(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.Adjusts(Mcols, Gcols, width);
            count++;
        }//调整连续多行列宽的方法
        private void AddData(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.AddData(Row, Column, Data);
            count++;
        }//添加数据
        private void AddFormula(string FilesPath)
        {
            ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath);
            ExcelProcess2.AddFormula (Column,Formula);
            count++;
        }//添加公式的方法
        private void WeatherTool(string FilesPath)
        {
            string [] a = FilesPath.Split('\\');
            string Fname = a[a.Length-1];
            string[] b = Fname.Split('.');
            string hz = b[1];
            string Filename = b[0];
            if (hz=="dly")
            {
                ExcelProcess2 = new ExcelProcess(FilesPath, ExportPath,Filename);
                ExcelProcess2.WeatherTool(param); ;
                count++;
            }

        }
        private void INPUT(string FilesPath)
        {
            string[] a = FilesPath.Split('\\');
            string Fname = a[a.Length - 1];
            string[] b = Fname.Split('.');
            string hz = b[1];
            string Filename = b[0];
            if (hz == "INP")
            {
                HZPath = ExportPath + @"\" + "INPUT汇总";
                ExcelProcess2.GatherINP(Filename, Span.ToString(), FilesPath);
                xlsWorkBook3 = ExcelProcess2.xlsWorkBook3;
                xlsWorkSheet3 = ExcelProcess2.xlsWorkSheet3;
                Span += 12;
                count++;
            }

        }//INP汇总
        private void  OUTPUT(string FilesPath)
        {
            string[] a = FilesPath.Split('\\');
            string Fname = a[a.Length - 1];
            string[] b = Fname.Split('.');
            string hz = b[1];
            string Filename = b[0];
            if (hz == "OUT")
            {
                HZPath = ExportPath + @"\" + "OUT汇总";
                 ExcelProcess2.GatherOUT(Filename, Span.ToString(), FilesPath);
                xlsWorkBook4 = ExcelProcess2.xlsWorkBook4;
                xlsWorkSheet4 = ExcelProcess2.xlsWorkSheet4;
                Span += 12;
                count++;
            }

        }//OUT汇总
        private void ProgressBar_Load(object sender, EventArgs e)
        {
            FileprogressBar.Maximum = (WdataFile.Length + 1) * 10;
            FileprogressBar.Value = 10;
            FileprogressBar.Step = 10;
        }
        private void SetValue(string wFileName)
        {
            this.FileLabel.Text = "程序正在运行" + wFileName;
            FileLabel.Refresh();
        }//显示正在处理的文件
        private void ChangeProgress()//进度条增长
        {
            FileprogressBar.Value += FileprogressBar.Step;
            Thread.Sleep(500);
            System.Windows.Forms.Application.DoEvents();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
            if (thread1!=null)
            {
                thread1.Abort();
                thread1.Join();
                GC.Collect();
            }
            if (ExcelProcess2!=null)
            {
                ExcelProcess2.KillSpecialExcel();
            }
          
        }
        private void  SaveTypes()
        {
            bool Clear = false;
            string wFileName = "";
            try
            {

                foreach (var Wfile in WdataFile)
                {
                    if (Stop)
                    {
                        wFileName = Wfile.Name;
                        FileLabel.Invoke(setCallBack, wFileName);
                        string FilesPath = ResourePath + @"\" + wFileName;
                        ExcelWay(FilesPath);
                        FileprogressBar.Invoke(setCallBack2);
                        //ExcelProcess2.KillSpecialExcel();
                    }
                }
                if (count != 0)
                {
                    MessageBox.Show("已处理全部" + count + "个文件,处理成功");
                }
                else
                {
                    MessageBox.Show("未处理文件，请检出资源文件夹是否正确");
                }
            }
            catch (ThreadAbortException)//线程中断
            {
                if (ExcelProcess2 != null)
                {
                    ExcelProcess2.KillSpecialExcel();
                }
            }
            catch (Exception e)//异常中断
            {
                MessageBox.Show("处理" + wFileName + " 文件时出现错误，请检查该文件是否符合处理条件");
                MessageBox.Show(e.ToString());
                if (ExcelProcess2 != null)
                {
                    ExcelProcess2.KillSpecialExcel();
                }
            }
            finally
            {
                if (ExcelProcess2 != null)
                {
                    ExcelProcess2.KillSpecialExcel();
                }
            }         
            if (xlsWorkBook3!=null)
            {
                xlsWorkBook3.SaveAs(HZPath);
                xlsWorkBook3.Close(false);
                ExcelProcess2.KillSpecialExcel();
            }
            if (xlsWorkBook4 != null)
            {
                xlsWorkBook4.SaveAs(HZPath);
                xlsWorkBook4.Close(false);
                ExcelProcess2.KillSpecialExcel();
            }
            
        }//遍历文件并应用不同的方法
        private void timer1_Tick(object sender, EventArgs e)//自动开启新线程
        {
            timer1.Enabled = false;
            thread1 = new Thread(new ThreadStart(SaveTypes));
            thread1.SetApartmentState(System.Threading.ApartmentState.STA);
            thread1.IsBackground = true;
            thread1.Start();//调用Start方法执行线程
        }
        private void CloseForm()
        {
            this.CloseForm();
            thread1.Abort();
            thread1.Join();
            GC.Collect();
        }
        private void ProgressBar_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (thread1 != null)
            {
                thread1.Abort();
                thread1.Join();
                GC.Collect();
            }
            if (ExcelProcess2 != null)
            {
                ExcelProcess2.KillSpecialExcel();
            }
        }
    }
    }

