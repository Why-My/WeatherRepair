using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections;
using System.Text.RegularExpressions;
namespace WeatherRepair
{
    class ExcelProcess
    {
        string Debug = System.Windows.Forms.Application.StartupPath;
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hwnd, out int processid);
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlsWorkBook;//王
        Excel.Workbook xlsWorkBook2;//BAOJI2.XLS
        Excel.Worksheet xlsWorkSheet;
        Excel.Worksheet xlsWorkSheet2;//国家气象局信息
        public static string StationCode;
        int colsCount;
        int column;
        string ExcPath;
        string ExcPath2;
        string ExposePath;
        string FileName;
        string ResoureFile;
        Excel.Application xlApp2;
        public  Excel.Workbook xlsWorkBook3;
        public Excel.Workbook xlsWorkBook4;
        public Excel.Worksheet xlsWorkSheet3;
        public Excel.Worksheet xlsWorkSheet4;
        int count;
        string StartRow;
        string EndRow;
        string HZPath;
        bool BAdd = true;
        public ExcelProcess(string path, string Exc)//初始化
        {
            string[] a = path.Split('\\');
            string Fname = a[a.Length - 1];
            string[] b = Fname.Split('.');
            string hz = b[1];
            FileName = b[0];
            ExposePath = Exc;
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            xlApp.Visible = false;
            xlApp.DisplayAlerts = true;
            xlsWorkBook = xlApp.Workbooks.Open(path);
            xlsWorkSheet = (Worksheet)xlsWorkBook.Worksheets[1];
            colsCount = xlsWorkSheet.UsedRange.Columns.Count;
            if (xlsWorkSheet.AutoFilterMode == true)
            {
                xlsWorkSheet.AutoFilterMode = false;
            }
            ExcPath = Exc + @"\" + xlsWorkSheet.Name;
            StationCode = ((Excel.Range)xlsWorkSheet.Cells[2, 1]).Text.ToString();
        }
        public ExcelProcess(string path,  string Exc,string Filename2)//初始化
        {
            ResoureFile = path;
            FileName = Filename2;
            ExposePath = Exc;
        }
        public ExcelProcess( )//初始化
        {
            xlApp2 = new Excel.Application();
            xlApp2.Visible = false;
           
        }
        public void SaveAs2(string SaveType)//另存文件
        {
            switch (SaveType)
            {
                case "气象数据常规格式(*.dly)":
                    ExcPath2 = ExcPath + ".prn";
                    FileExist(ExcPath2);
                    xlsWorkBook.SaveAs(ExcPath, Excel.XlFileFormat.xlTextPrinter);
                    xlsWorkBook.Close(false);
                    string prnPath = ExposePath + @"\" + FileName + ".prn";
                    File.Move(ExposePath + @"\" + FileName + ".prn", ExposePath + @"\" + FileName + ".dly");
                    File.Delete(prnPath);
                    break;
                case "文本文件(制表符分隔)(*.txt)":
                    ExcPath2 = ExcPath + ".txt";
                    FileExist(ExcPath2);
                    //xlsWorkBook.SaveAs(ExcPath, Excel.XlFileFormat.xlTextWindows);
                    xlsWorkBook.SaveAs(ExcPath, Excel.XlFileFormat.xlUnicodeText);
                    xlsWorkBook.Close(false);
                    break;
                case "带格式文本文件(空格分隔)(*.prn)":
                    ExcPath2 = ExcPath + ".prn";
                    FileExist(ExcPath2);
                    xlsWorkBook.SaveAs(ExcPath, Excel.XlFileFormat.xlTextPrinter);
                    xlsWorkBook.Close(false);
                    break;
                default:
                    break;
            }
        }
        public void FileExist(string ExcPath2)//判断文件是否已存在
        {
            if (File.Exists(ExcPath2))
            {
                File.Delete(ExcPath2);
            }


        }
        public void KillSpecialExcel()
        {
            try
            {
                if (xlApp != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out lpdwProcessId);
                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Delete Excel Process Error:" + ex.Message);
            }
        }
        #region Excel非连续空缺处理
        public void SingleRow()
        {
            char E = 'A';
            int column = 1;
            int e = Convert.ToInt32(E);
            while (column <= colsCount)
            {
                CountName = ((char)(e++)).ToString();//选定列号
                SingleConfirmCols(CountName, column);
                column++;
            }
            ((Excel.Range)xlsWorkSheet.Cells[1, 12]).EntireColumn.Delete(0);
            SaveAs();

        }
        public void SingleConfirmCols(string cols, int column)
        {
            try
            {
                Range Mange2 = xlsWorkSheet.get_Range(cols + "1", cols + xlsWorkSheet.UsedRange.Rows.Count).SpecialCells(Excel.XlCellType.xlCellTypeBlanks);

                int length = Mange2.Count;
                int[] Rowlist = new int[length];
                int j = 0;
                foreach (Excel.Range r in Mange2)
                {
                    int RowCount = r.Row;
                    Rowlist[j] = RowCount;
                    j++;
                }

                SingleDay(Rowlist, cols, column);

            }
            catch (Exception)
            {


            }




        }//对单个空缺进行操作
        public void SingleDay(int[] Rowlist, string cols, int column)//修补数据
        {
            int PDay = Rowlist[0];
            int Year1 = Convert.ToInt32(((Range)xlsWorkSheet.Cells[PDay, 2]).Text.ToString());
            int Month1 = Convert.ToInt32(((Range)xlsWorkSheet.Cells[PDay, 3]).Text.ToString());
            double Average = ThreeAverage(Year1, Month1, cols);
            xlsWorkSheet.Cells[PDay, column] = Average.ToString();

            for (int i = 0; i < Rowlist.Length - 1; i++)
            {

                int QDay = Rowlist[i + 1];
                int Year2 = Convert.ToInt32(((Range)xlsWorkSheet.Cells[QDay, 2]).Text.ToString());
                int Month2 = Convert.ToInt32(((Range)xlsWorkSheet.Cells[QDay, 3]).Text.ToString());

                if (Year1 != Year2 || Month1 != Month2)
                {
                    double Average2 = ThreeAverage(Year2, Month2, cols);
                    xlsWorkSheet.Cells[QDay, column] = Average2.ToString();
                }
                else
                {
                    xlsWorkSheet.Cells[QDay, column] = Average.ToString();

                }
                Year1 = Year2;
                Month1 = Month2;
            }



        }
        public double ThreeAverage(int Year, int Month, string Cols)
        {
            int exam = Convert.ToInt32(((Range)xlsWorkSheet.Cells[2, 2]).Text.ToString());
            int Year1 = Year - 1;
            int Year2 = Year - 2;
            int Year3 = Year - 3;

            if (exam > Year3)
            {
                Year1 = Year + 1;
                Year2 = Year + 2;
                Year3 = Year + 3;
                string Formal = string.Format("=ROUND(SUBTOTAL(101,{0}:{1}),2)", Cols, Cols);
                xlsWorkSheet.get_Range("B1", Type.Missing).AutoFilter(2, Year1.ToString(), XlAutoFilterOperator.xlOr, Year2.ToString(), Type.Missing);
                xlsWorkSheet.get_Range("C1", Type.Missing).AutoFilter(3, Month.ToString(), XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);
                xlsWorkSheet.Cells[1, 12] = Formal;//R
                double Average1 = Convert.ToDouble(((Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[1, 12]).Text.ToString());
                xlsWorkSheet.get_Range("B1", Type.Missing).AutoFilter(2, Year3.ToString(), XlAutoFilterOperator.xlOr, Type.Missing, Type.Missing);
                xlsWorkSheet.Cells[2, 12] = Formal;//R
                double Average2 = Convert.ToDouble(((Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[2, 12]).Text.ToString());
                double Average3 = (Average1 + Average2) / 2;
                double Average = Math.Round(Average3, 1);
                xlsWorkSheet.AutoFilterMode = false;
                return Average;
            }
            else
            {
                string Formal = string.Format("=ROUND(SUBTOTAL(101,{0}:{1}),2)", Cols, Cols);
                xlsWorkSheet.get_Range("B1", Type.Missing).AutoFilter(2, Year1.ToString(), XlAutoFilterOperator.xlOr, Year2.ToString(), Type.Missing);
                xlsWorkSheet.get_Range("C1", Type.Missing).AutoFilter(3, Month.ToString(), XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);
                xlsWorkSheet.Cells[1, 12] = Formal;//R
                double Average1 = Convert.ToDouble(((Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[1, 12]).Text.ToString());
                xlsWorkSheet.get_Range("B1", Type.Missing).AutoFilter(2, Year3.ToString(), XlAutoFilterOperator.xlOr, Type.Missing, Type.Missing);
                xlsWorkSheet.Cells[2, 12] = Formal;//R
                double Average2 = Convert.ToDouble(((Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[2, 12]).Text.ToString());
                double Average3 = (Average1 + Average2) / 2;
                double Average = Math.Round(Average3, 1);
                xlsWorkSheet.AutoFilterMode = false;
                return Average;

            }


        }//求平均数
        #endregion
        #region Excel连续空缺处理
        string CountName;
        public void SuccessionRow(string Days2)//处理连续空缺天数大于十天的数据
        {

            int Days = Convert.ToInt32(Days2);
            char E = 'A';
            int column = 1;
            int e = Convert.ToInt32(E);
            while (column++ <= colsCount)
            {
                CountName = ((char)(e++)).ToString();//选定列号
                ConfirmCols(CountName, Days);

            }
            SaveAs();
            //MessageBox.Show("OK");


        }
        public void RepairData(int PDay, int Sum)//修补数据
        {
            int FDay = PDay - Sum;
            int exam = Convert.ToInt32(((Range)xlsWorkSheet.Cells[2, 2]).Text.ToString());
            int Year = Convert.ToInt32(((Range)xlsWorkSheet.Cells[PDay, 2]).Text.ToString());
            int FMonth = Convert.ToInt32(((Range)xlsWorkSheet.Cells[FDay, 3]).Text.ToString());
            int LMonth = Convert.ToInt32(((Range)xlsWorkSheet.Cells[PDay, 3]).Text.ToString());
            bool Ryear = DateTime.IsLeapYear(Year);


            if (Year == exam)
            {
                bool Byear = DateTime.IsLeapYear(Year + 1);
                int XDays = MinusDay2(Ryear, Byear, FMonth, LMonth);
                if (XDays == 2)
                {
                    CutDays2(Year, LMonth, FDay, PDay, XDays, 366, 365);


                }
                else if (XDays == 1)
                {
                    CutDays2(Year, LMonth, FDay, PDay, XDays, 365, 366);
                }
                else
                {
                    int PDay2 = PDay + XDays;
                    int FDay2 = PDay2 - Sum;
                    Paste(FDay, PDay, FDay2, PDay2, CountName);
                }
            }
            else
            {
                bool Byear = DateTime.IsLeapYear(Year - 1);
                int XDays = MinusDay(Ryear, Byear, FMonth, LMonth);
                if (XDays == 2)
                {
                    CutDays(Year, LMonth, FDay, PDay, XDays, 365, 366);


                }
                else if (XDays == 1)
                {
                    CutDays(Year, LMonth, FDay, PDay, XDays, 366, 365);
                }
                else
                {
                    int PDay2 = PDay - XDays;
                    int FDay2 = PDay2 - Sum;
                    Paste(FDay, PDay, FDay2, PDay2, CountName);
                }
            }
        }
        public int MinusDay(bool Ryear, bool Byear, int FMonth, int LMonth)//获取间隔天数
        {

            if (Ryear)
            {
                if (FMonth > 2)
                {
                    return 366;
                }
                else if (LMonth < 3)
                {
                    return 365;
                }
                else if (FMonth < 3 && LMonth > 2)
                {
                    return 2;
                }

            }
            else if (Byear)
            {
                if (FMonth > 2)
                {
                    return 365;
                }
                else if (LMonth < 3)
                {
                    return 366;
                }
                else if (FMonth < 3 && LMonth > 2)
                {
                    return 1;
                }
            }


            return 365;



        }
        public int MinusDay2(bool Ryear, bool Nextyear, int FMonth, int LMonth)
        {
            if (Ryear)
            {
                if (FMonth > 2)
                {
                    return 365;
                }
                else if (LMonth < 3)
                {
                    return 366;
                }
                else if (FMonth < 3 && LMonth > 2)
                {
                    return 2;
                }

            }
            else if (Nextyear)
            {
                if (FMonth > 2)
                {
                    return 366;
                }
                else if (LMonth < 3)
                {
                    return 365;
                }
                else if (FMonth < 3 && LMonth > 2)
                {
                    return 1;
                }
            }


            return 365;






        }
        public void Paste(int FDay1, int LDay1, int FDay2, int Lday2, string cols)//粘贴数据
        {


            Range XDatarange = xlsWorkSheet.get_Range(cols + FDay2, cols + Lday2);
            XDatarange.Copy();
            Range Datarange = xlsWorkSheet.get_Range(cols + FDay1, cols + LDay1);
            Datarange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


        }
        public void CutDays(int Year, int LMonth, int FDay, int PDay, int increase, int Fdata, int Ldata)//将数据分段
        {
            int LDay = Convert.ToInt32(((Range)xlsWorkSheet.Cells[PDay, 4]).Text.ToString());
            string SData = Year + "-" + LMonth + "-" + LDay;
            string February = Year + "-2-28";
            DateTime dt1 = Convert.ToDateTime(February);
            DateTime dt2 = Convert.ToDateTime(SData);
            TimeSpan span = dt2.Subtract(dt1);
            int TwentyEight = PDay - span.Days;
            int ThirtyOne = TwentyEight + increase;
            int TwentyEight2 = TwentyEight - Fdata;
            int ThirtyOne2 = ThirtyOne - Ldata;
            int FDay3 = FDay - Fdata;
            int PDay3 = PDay - Ldata;
            Paste(FDay, TwentyEight, FDay3, TwentyEight2, CountName);
            Paste(ThirtyOne, PDay, ThirtyOne2, PDay3, CountName);

        }
        public void CutDays2(int Year, int LMonth, int FDay, int PDay, int increase, int Fdata, int Ldata)//将数据分段
        {
            int LDay = Convert.ToInt32(((Range)xlsWorkSheet.Cells[PDay, 4]).Text.ToString());
            string SData = Year + "-" + LMonth + "-" + LDay;
            string February = Year + "-2-28";
            DateTime dt1 = Convert.ToDateTime(February);
            DateTime dt2 = Convert.ToDateTime(SData);
            TimeSpan span = dt2.Subtract(dt1);
            int TwentyEight = PDay - span.Days;
            int ThirtyOne = TwentyEight + increase;
            int TwentyEight2 = TwentyEight + Fdata;
            int ThirtyOne2 = ThirtyOne + Ldata;
            int FDay3 = FDay + Fdata;
            int PDay3 = PDay + Ldata;
            Paste(FDay, TwentyEight, FDay3, TwentyEight2, CountName);
            Paste(ThirtyOne, PDay, ThirtyOne2, PDay3, CountName);

        }
        public void ConfirmCols(string cols, int Days)
        {

            try
            {
                Range Mange2 = xlsWorkSheet.get_Range(cols + "1", cols + xlsWorkSheet.UsedRange.Rows.Count).SpecialCells(Excel.XlCellType.xlCellTypeBlanks);

                int length = Mange2.Count;
                int[] Rowlist = new int[length];
                int j = 0;
                foreach (Excel.Range r in Mange2)
                {
                    int RowCount = r.Row;
                    Rowlist[j] = RowCount;
                    j++;
                }
                int Sum = 0;
                for (int i = 0; i < Rowlist.Length - 1; i++)
                {
                    int PDay = Rowlist[i];
                    int QDay = Rowlist[i + 1];
                    if ((QDay - PDay) == 1)
                    {
                        Sum++;
                    }
                    else if ((QDay - PDay) != 1 && Sum >= Days)
                    {
                        RepairData(PDay, Sum);
                        Sum = 0;
                    }
                    else if ((QDay - PDay) != 1 && Sum < Days)
                    {
                        Sum = 0;
                    }

                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {


            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());

            }

        }//对选定的列进行操作
        #endregion
        #region 气象数据格式调整
        public void MoveClos(string Mcols, string Gcols)
        {

            Range ZRRange = xlsWorkSheet.get_Range(Mcols + "1", Mcols + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            ZRRange.Cut();
            Range MaxRange = xlsWorkSheet.get_Range(Gcols + "1", Gcols + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            MaxRange.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            xlsWorkBook.Close(true);


        }//移动列
        public void DeleteRows(string Mcols, string Gcols)//删除多行
        {
            int M = Convert.ToInt32(Mcols);
            int G = Convert.ToInt32(Gcols);
            Range D1_rng = xlsWorkSheet.Range[xlsWorkSheet.Cells[M, 1], xlsWorkSheet.Cells[G, colsCount]];
            D1_rng.Delete();
            xlsWorkBook.Close(true);

        }
        public void DeleteRow(string Scols)//删除单行
        {
            int S = Convert.ToInt32(Scols);
            Range deleteRng = (Range)xlsWorkSheet.Rows[S, System.Type.Missing];
            deleteRng.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            xlsWorkBook.Close(true);

        }
        public void DeleteCols(string Mcols, string Gcols)//删除连续多列
        {
            Range xDeleteRange2 = xlsWorkSheet.get_Range(Mcols + "1", Gcols + xlsWorkSheet.UsedRange.Rows.Count);
            xDeleteRange2.Delete();
            xlsWorkBook.Close(true);

        }
        public void DeleteFormulas(string Mcols, string Gcols)//删除连续多列
        {
            Range xDeleteRange2 = xlsWorkSheet.get_Range(Mcols + "1", Gcols + xlsWorkSheet.UsedRange.Rows.Count);
            xDeleteRange2.Copy();
            xDeleteRange2.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Clipboard.Clear();
            xlsWorkBook.Close(true);

        }
        public void DeleteCol(string Scols)//删除单列
        {

            Range ZRRange = xlsWorkSheet.get_Range(Scols + "1", Scols + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            ZRRange.Delete();
            xlsWorkBook.Close(true);

        }
        public void ChangeType(string Scols)//删除数据格式
        {
            int a = xlsWorkSheet.UsedRange.Columns.Count + 1;
            xlsWorkSheet.Cells[1, a] = 1;
            int b = a + 1;
            char c = (char)(b + 64);
            char d = (char)(a + 64);
            string Column = c.ToString();
            string Column2 = d.ToString();

            Range ScolsRange = xlsWorkSheet.get_Range(Scols + "1", Scols + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            Range ZRRange = xlsWorkSheet.get_Range(Column + "1", Column + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            string Formula2 = string.Format(@"={0}1*${1}$1", Scols, Column2);
            ZRRange.Formula = Formula2;
            ZRRange.Copy();
            ScolsRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            ZRRange.Delete();
            xlsWorkSheet.Cells[1, a] = "";
            xlsWorkBook.Close(true);

        }
        public void DeleteFormula(string Scols)
        {
            //Thread th = Thread.CurrentThread;
            //th.SetApartmentState(System.Threading.ApartmentState.STA);
            Range ZRRange = xlsWorkSheet.get_Range(Scols + "1", Scols + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            ZRRange.Copy();
            ZRRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Clipboard.Clear();
            xlsWorkBook.Close(true);

        }//清空单列公式
        public void Adjust(string Scols, int width)
        {

            Range ZRRange = xlsWorkSheet.get_Range(Scols + "1", Scols + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            ZRRange.ColumnWidth = width;
            xlsWorkBook.Close(true);

        }//调整列宽
        public void Adjusts(string Mcols, string Gcols, int width)
        {

            Range xDeleteRange2 = xlsWorkSheet.get_Range(Mcols + "1", Gcols + xlsWorkSheet.UsedRange.Rows.Count);
            xDeleteRange2.ColumnWidth = width;

            xlsWorkBook.Close(true);
        }//调整多列列宽
        public void AddData(string Row, string Column2, string Data)
        {
           
            string a = Column2;
            char[] A = a.ToCharArray();
            int AA = A[0];
            if (AA>64&&AA<91)
            {
                column = AA - 64;
            }
            if (AA > 96 && AA < 123)
            {
                column = AA -96 ;
            }
            //MessageBox.Show(AA.ToString());
            int row = Convert.ToInt32(Row);
            xlsWorkSheet.Cells[row, column] = Data;
            xlsWorkBook.Close(true);
        }//添加数据
        public void AddFormula(string Cloumn, string Formula2)
        {
            Range xStationCode = xlsWorkSheet.get_Range(Cloumn+"1", Cloumn + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            string Formula = string.Format(@"={0}", Formula2);
           // MessageBox.Show(Formula);
            try
            {
                xStationCode.Formula = Formula;
            }
            catch (Exception e)
            {

                MessageBox.Show(e.ToString());
            }
           
            xlsWorkBook.Close(true);
        }//添加公式
        public void SaveAs()
        {
            ExcPath2 = ExcPath + ".xlsx";
            FileExist(ExcPath2);
            ExcPath2 = ExcPath + ".xls";
            FileExist(ExcPath2);
            xlsWorkBook.SaveAs(ExcPath+"2", Excel.XlFileFormat.xlWorkbookDefault);

        }
        #endregion
        #region 计算逐日太阳辐射量
        public void ZSSunlight(string LatPath, string Sunglight)//获取逐日太阳辐射量
        {
            string F = Sunglight;
            int a = xlsWorkSheet.UsedRange.Columns.Count + 1;
            xlsWorkSheet.Cells[1, a] = 1;
            int b = (a + 64);
            char l = (char)b;
            char m = (char)(b+1);
            char n = (char)(b + 2);
            char o = (char)(b + 3);
            char p = (char)(b + 4);
            char q = (char)(b + 5);
            char r = (char)(b + 6);

            string L = l.ToString();//1
            string M = m.ToString();
            string N = n.ToString();
            string O = o.ToString();
            string P = p.ToString();
            string Q = q.ToString();
            string R = r.ToString();

            string LFormula = string.Format(@"=IF({0}1+1>IF(INT(B1/4)=B1/4,366,365),1,{0}1+1)",L);
            string MFormula = string.Format(@"=SIN(6.28/365*({0}1-80.25))*0.4102", L);
            string OFormula = string.Format(@"=ACOS(-TAN(6.28/360*{0}1)*TAN({1}1))",N,M);
            string PFormula = string.Format(@"=30*(1+0.0335*SIN(6.28/365*({0}1+88.2)))*(O1*SIN(6.28/360*{1}1)*SIN({2}1)+COS(6.28/360*{1}1)*COS({2}1)*SIN({3}1))", L, N,M,O);
            string QFormula = string.Format(@"=7.64*ACOS((-SIN(6.28/360*{0}1)*SIN({1}1)-0.044)/COS(6.28/360*{0}1)*COS({1}1))", N, M);
            string RFormula = string.Format(@"=ROUND({0}1*(0.18+0.55*({1}1/{2}1)),2)", P, F,Q);

            Range LRang = xlsWorkSheet.get_Range(L+"2", L + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            LRang.Formula = LFormula;
            Range MRang = xlsWorkSheet.get_Range(M+"1", M + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            MRang.Formula = MFormula;
            Lat(LatPath,N);
            Range ORang = xlsWorkSheet.get_Range(O+"1", O + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            ORang.Formula = OFormula;
            Range PRang = xlsWorkSheet.get_Range(P+"1", P + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            PRang.Formula = PFormula;
            Range QRang = xlsWorkSheet.get_Range(Q+"1", Q + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            QRang.Formula = QFormula;
            Range RRang = xlsWorkSheet.get_Range(R+"1", R + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            RRang.Formula = RFormula;
            RRang.Copy();
            RRang.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Range xDeleteRange2 = xlsWorkSheet.get_Range(L+"1", Q + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            xDeleteRange2.Delete();
             Clipboard.Clear();
            xlsWorkBook.Close(true);
        }
        public void Lat(string LatPath,string N)//获取lat
        {
            string [] lat = LatPath.Split('\\');
            int a = lat.Length - 1;
            string lat2 = lat[a];
            xlsWorkBook2 = xlApp.Workbooks.Open(LatPath);
            xlsWorkSheet2 = (Worksheet)xlsWorkBook2.Worksheets[1];
            Range Latrng = xlsWorkSheet.get_Range(N+"1", N + xlsWorkSheet.UsedRange.Rows.Count);//设置列范围
            string LFormula = string.Format(@"=LOOKUP(A1,[{0}]Sheet1!$A:$A,[{0}]Sheet1!$B:$B)", lat2);
            Latrng.Formula = LFormula;


        }
        #endregion
        #region WeatherTool
        public void WeatherTool(string Param)
        {
           
            MoveFile();
            FileWrite(Param);
            System.Diagnostics.Process.Start(Debug + @"\WXPM3020.exe");//执行程序
            Thread.Sleep(2000);
            FileMove();
        }
        public void MoveFile()
        {
           // string a = ResoureFile;
            File.Copy(ResoureFile , Debug + @"\" + FileName + ".dly");

        }
        public void FileWrite(string Param)//写入文件
        {
            FileStream stream = new FileStream(Debug + @"\" + FileName + ".dly", FileMode.Open);//fileMode指定是读取还是写入
            StreamReader read = new StreamReader(stream);
            string a = read.ReadToEnd();
            read.Close();
            stream.Close();//释放内存
            FileStream stream2 = new FileStream(Debug + @"\" + FileName + ".dly", FileMode.Create);//fileMode指定是读取还是写入
            StreamWriter write = new StreamWriter(stream2);
            write.WriteLine(FileName);
            write.WriteLine("  "+ Param);
           // write.WriteLine("  361980");
            write.WriteLine(a);
            write.Close();
            stream2.Close();

            FileStream stream3 = new FileStream(Debug + @"\WXPMRUN.DAT", FileMode.Create);//fileMode指定是读取还是写入
            StreamWriter write3 = new StreamWriter(stream3);
            write3.WriteLine(FileName + ".dly");
            write3.WriteLine();
            write3.WriteLine();
            write3.Close();
            stream3.Close();
        }
        public void FileMove()//移动文件
        {
            File.Move(Debug + @"\" + FileName + ".INP", ExposePath + @"\inp\" + FileName + ".INP");
            File.Move(Debug + @"\" + FileName + ".OUT", ExposePath + @"\out\" + FileName + ".OUT");
            File.Move(Debug + @"\" + FileName + ".dly", ExposePath + @"\Changedly\" + FileName + ".dly");


        }
        #endregion
        #region input/output汇总
        public void GatherINP(string Filename2, string count2,string FilesPath)

        {
            int span = Convert.ToInt32(count2);
            
            if (File.Exists(HZPath))
            {
                MessageBox.Show("汇总文件已存在，请删除");
            }
            else
            {
                if (BAdd)
                {
                    xlsWorkBook3 = xlApp2.Workbooks.Add(true);
                    BAdd = false;
                }
                xlsWorkSheet3 = xlsWorkBook3.ActiveSheet;
                xlsWorkSheet3.Name = "input汇总";

            }
            double[,] input2 = InpExcel(FilesPath);
            XInputExcel(Filename2, input2, span);
           
        }
        public void GatherOUT(string Filename2, string count2, string FilesPath)
        {
            int span = Convert.ToInt32(count2);
            if (File.Exists(HZPath))
            {
                MessageBox.Show("汇总文件已存在，请删除");
            }
            else
            {
                if (BAdd)
                {
                    xlsWorkBook4 = xlApp2.Workbooks.Add(true);
                    BAdd = false;
                }
              
                xlsWorkSheet4 = xlsWorkBook4.ActiveSheet;
                xlsWorkSheet4.Name = "WSPD汇总";

            }
            double[] outPut = OutExcel(FilesPath);
            XOutputExcel(Filename2, outPut, span);
            
        }
        public double[,] InpExcel(string FilesPath)//inp导入excel
        {
            //将INP文件读入数组

            string[] lines = System.IO.File.ReadAllLines(FilesPath);
            ArrayList input = new ArrayList();
            int position = 0;
            for (int i = 2; i < lines.Length; i++)
            {
                string row = lines[i];

                while (position <= row.Length - 6)
                {
                    input.Add(row.Substring(position, 6));
                    position += 6;

                }
                position = 0;
            }
            double[,] input2 = new double[12, 12];
            int j = 0;
            foreach (var item in input)
            {
                input2[j / 12, j % 12] = Convert.ToDouble(item);
                j++;
            }
            return input2;
        }
        public double[] OutExcel(string FilesPath) //out导入excel
        {
            FileStream stream = new FileStream(FilesPath, FileMode.Open);//fileMode指定是读取还是写入
            StreamReader read = new StreamReader(stream);
            string sr = read.ReadToEnd();
            int wspd1 = sr.IndexOf("WSPD");
            int wspd2 = sr.LastIndexOf("WSPD"); ;
            string WSPD = sr.Substring(wspd1 + 9, wspd2 - wspd1 - 11);
            read.Close();
            stream.Close();//释放内
            string Out = Regex.Replace(WSPD, "\\s{2,}", " ");
            string[] outexcel = Out.Split(' ');
            double[] outexcel2 = Array.ConvertAll<string, double>(outexcel, value => Convert.ToDouble(value));
            return outexcel2;
        }
        public void XInputExcel(string FileName, double[,] input, int count)
        {
            StartRow = count.ToString();
            EndRow = (count + 11).ToString();
            string[] Title = new string[12] { "TMX", "TMN", "SDMX", "SDMN", "PRCP", "SDRF", "SKRF", "PW|D", "PW|W", "DAYP", "RAD", "RHUM" };
            xlsWorkSheet3.get_Range("A" + StartRow, "A" + EndRow).Resize.Value2 = FileName;
            xlsWorkSheet3.get_Range("B" + StartRow, "B" + EndRow).Value2 = xlApp2.WorksheetFunction.Transpose(Title);
            xlsWorkSheet3.get_Range("C" + StartRow, "N" + EndRow).Resize.Value2 = input;

        }
        public void XOutputExcel(string FileName, double[] outPut2, int count)
        {
            StartRow = count.ToString();
            EndRow = (count + 11).ToString();
            int[] Month = new int[12] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            xlsWorkSheet4.get_Range("A" + StartRow, "A" + EndRow).Resize.Value2 = FileName;
            xlsWorkSheet4.get_Range("B" + StartRow, "B" + EndRow).Value2 = xlApp2.WorksheetFunction.Transpose(Month);
            xlsWorkSheet4.get_Range("C" + StartRow, "C" + EndRow).Value2 = xlApp2.WorksheetFunction.Transpose(outPut2);
        }
        #endregion
    }
}