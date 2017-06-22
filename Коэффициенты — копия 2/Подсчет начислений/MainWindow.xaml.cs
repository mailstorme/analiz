using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using win = System.Windows;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Web;
using PdfSharp.Pdf.Printing;
using System.Diagnostics;

//using System.mscorlib;

namespace Подсчет_начислений
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string[] R1C1 = new string[] { "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ" };


        public MainWindow()
        {
            InitializeComponent();
        }


        public void CloseProcess(Process[] before)
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                if (!before.Contains(proc))
                    proc.Kill();
            }
        }


        private object[][] getarray(string path,int[] columns /*,int c1, int c2,int c3,int tarif*/)
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            int Rows = excelworksheet.UsedRange.Rows.Count;
            int Columns = excelworksheet.UsedRange.Columns.Count;

            object[][] arr = new object[4][];

            int icolumn = 0;
            foreach(int column in columns)
            {
                for (int i = 0; i < Columns + 1; i++)
                {
                    if (column == i)
                    {
                        object[,] massiv;
                        arr[icolumn] = new object[Rows - 1];
                        range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                        massiv = (System.Object[,])range.get_Value(Type.Missing);
                        arr[icolumn] = massiv.Cast<object>().ToArray();
                        icolumn++;
                    }
                }
            }

            #region прошлый вариант взятия колонок
            /*
            int icount = 0;
            for (int i = 0; i < Columns + 1; i++)
                if (i == c1 || i == c2 || i == c3)
                {
                    object[,] massiv;
                    arr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    arr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }
            for (int i = 0; i < Columns + 1; i++)
                if (i == tarif)
                {
                    object[,] massiv;
                    arr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    arr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }
                */
#endregion

            #region Закрытие Excel

            book.Close(false,false,false);

            ExcelApp.Quit();

            
            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);
            

            return arr;
        }


        private object[][] getbasearray(string path)
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            int Rows = excelworksheet.UsedRange.Rows.Count;
            int Columns = excelworksheet.UsedRange.Columns.Count;

            object[][] arr = new object[2][];

            int icount = 0;
            for (int i = 0; i < Columns + 1; i++)
                if (i == 1 || i == 2)
                {
                    object[,] massiv;
                    arr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    arr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }

            #region Закрытие Excel

            book.Close(false, false, false);

            ExcelApp.Quit();

            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            //workbooks = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);

            return arr;
        }


        #region Выбор папки
        private string DirSelect()
        {
            FolderBrowserDialog DirDialog = new FolderBrowserDialog();
            DirDialog.Description = "Выбор директории";
            DirDialog.SelectedPath = @"C:\";
            DirDialog.ShowDialog();
            return DirDialog.SelectedPath;
        }
        #endregion


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int _icc, _msisdn, _date;
            try
            {
                _icc = Convert.ToInt32(icc.Text, 10);
                _msisdn = Convert.ToInt32(msisdn.Text, 10);
                _date = Convert.ToInt32(date.Text, 10);
            }
            catch
            {
                win.MessageBox.Show("Ошибка входных данных");
                return;
            }

            string wpath = DirSelect();
            if (wpath == null || wpath == "")
            {
                win.MessageBox.Show("Выберите папку");
                return;
            }
            string[] files = Directory.GetFiles(wpath);


            string ComisPath = new OpenExcelFile().Filenamereturn();
            if (ComisPath == "can not open file")
                return;

            int Rows;
            int Columns;

            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range;

            book = ExcelApp.Workbooks.Open(ComisPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            Rows = excelworksheet.UsedRange.Rows.Count;
            Columns = excelworksheet.UsedRange.Columns.Count;

            object[][] ComisAr = new object[4][];
            //win.MessageBox.Show("Row=" + Rows + "  Column="+Columns);

            

            int icount = 0;
            for (int i = 0; i < Columns + 1; i++)
                if (i == _icc || i == _msisdn || i == _date )
                {
                    object[,] massiv;
                    ComisAr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    ComisAr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }

            ComisAr[icount] = new object[Rows - 1];
            for (int i = 0; i < Rows - 1; i++)
            {
                ComisAr[icount][i] = (double)0;
            }

            double sum = 0;
            int counts = 0;
            foreach (string file in files)
            {
                object[][] refill = getarray(file,new int[] { 4,5,6});
                int N = refill[1].Length;

                object sum1 = 0;
                for (int i = 0; i < N; i++)
                {
                    sum1 = Convert.ToDouble(sum1) + Convert.ToDouble(refill[1][i]);
                }

                win.MessageBox.Show(sum1.ToString());
               

                for (int i = 0; i < N; i++)
                {
                    int n = Array.BinarySearch(ComisAr[1], refill[2][i]) ;
                    if (n > 0)
                    {
                        counts++;
                        //double h1 = (double)ComisAr[icount][n];
                        //double h2 = (double)refill[1][i];
                        ComisAr[icount][n] = Convert.ToDouble(ComisAr[icount][n]) + Convert.ToDouble(refill[1][i]);
                        sum += (double)refill[1][i];
                    }
                }
            }

            object[,] arr = new object[Rows - 1, 1];
            //double qwe = 0;
            for(int i = 0; i < Rows-1; i++)
            {
                arr[i, 0] = ComisAr[icount][i];
                //qwe += Convert.ToDouble(ComisAr[icount][i]);
            }

            range = null;
            range = excelworksheet.get_Range(R1C1[Columns+1] + "2:" + R1C1[Columns + 1] + Rows);
            range.Value2 = arr;

            #region Закрытие Excel
            ExcelApp.Application.Quit();
            #endregion
            CloseProcess(List);


            win.MessageBox.Show("|" + sum.ToString() + "| в " + counts.ToString());
            GC.Collect();
        }

        private void msisdn_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CultureInfo culture;

            culture = CultureInfo.CurrentCulture;
            try
            {
                int number = Int32.Parse(e.Text, culture.NumberFormat);
            }
            catch (FormatException)
            {
                e.Handled = true;
            }
        }


        






        class diler
        {
            public object name;
            
            public int b;
            public int a;
            public int allincom;

            public int count1201;
            public int count1202;
            public int count1203;
            public int count12046;
            public int count120712;

            public int Tab;
            public int TabAll;
            public int Treg;
            public int TregAll;


            public double sum;

            public diler (object NAME, bool fir, bool sec, bool thi, bool from4, bool from7,object nachislenia,bool abonent,bool regula, bool abonentAll, bool regulaAll)
            {
                name = NAME;

                count1201 = 0;
                count1202 = 0;
                count1203 = 0;
                count12046 = 0;
                count120712 = 0;

                if (fir)
                    count1201 = 1;
                if (sec)
                    count1202 = 1;
                if (thi)
                    count1203 = 1;
                if(from4)
                    count12046 = 1;
                if (from7)
                    count120712 = 1;

                b = 0;
                a = 0;
                Tab = 0;
                TabAll = 0;
                Treg = 0;
                TregAll = 0;

                if (abonent)
                    Tab++;
                else if (regula)
                    Treg++;

                if (abonentAll)
                    TabAll++;
                else if (regulaAll)
                    TregAll++;

                allincom = 1;
                sum += Convert.ToDouble(nachislenia);

            }
        }


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string abonents = "МегаФон - Всё включено L МегаФон - Всё включено L 07.16 МегаФон - Всё включено M МегаФон - Всё включено M 07.16 МегаФон - Всё включено M+ МегаФон - Всё включено S МегаФон - Всё включено S 07.16 МегаФон - Всё включено S+ МегаФон - Всё включено VIP МегаФон - Всё включено XL МегаФон - Всё включено L 12.16 МегаФон - Всё включено M 12.16 МегаФон - Всё включено S 12.16 МегаФон - Всё включено VIP 07.16 МегаФон - Всё включено VIP 12.16 МегаФон - Всё включено XL 12.16 МегаФон - Всё включено XL 07.16 МегаФон.Безлимит МегаФон Онлайн с роутером 4G +" +
" МегаФон - Онлайн МегаФон - Онлайн с модемом 4G +   Связь городов   Тёплый приём   Тёплый приём +   Тёплый приём 2017 Тёплый приём M Тёплый приём M 2017 Тёплый приём M v.06.16 Тёплый приём S Тёплый приём S 2017 Тёплый приём S v.06.16";
            string regular = "МегаФон - Умный дом Переходи на НОЛЬ Переходи на НОЛЬ 2013";
            string NotFound = "";


            //string path1 = @"C:\Users\Andrei\Desktop\абонентская.txt";  мега
            //string path2 = @"C:\Users\Andrei\Desktop\регулярная.txt";

            string path1 = @"C:\Users\Andrei\Desktop\МтсАбонент.txt";
            string path2 = @"C:\Users\Andrei\Desktop\МтсРегуляр.txt";


            abonents = System.IO.File.ReadAllText(path1).Replace("\n", " ");
            regular = System.IO.File.ReadAllText(path2).Replace("\n", " ");

            //win.MessageBox.Show(abonents);
            //win.MessageBox.Show(regular);
            //return;

            string ComisPath = new OpenExcelFile().Filenamereturn();
            if (ComisPath == "can not open file")
                return;

            int Rows;
            int Columns;

            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range;

            book = ExcelApp.Workbooks.Open(ComisPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion

            Rows = excelworksheet.UsedRange.Rows.Count;
            Columns = excelworksheet.UsedRange.Columns.Count;

            object[,] ComisAr = new object[Rows,Columns + 11];

            range = excelworksheet.get_Range(R1C1[1] + "1:" + R1C1[Columns] + Rows.ToString());
            ComisAr = (System.Object[,])range.get_Value(Type.Missing);


            string mas = "";
            for (int i = 1; i <= 10; i++)
            {
                for (int j = 1; j <= Columns; j++)
                {
                    mas += ComisAr[i, j] + " ";
                }
                mas += "\n";
            }


            List<diler> dilers = new List<diler>();
            
            for (int i = 2; i <= Rows; i++)
            {
                if (ComisAr[i, 1] == null || ComisAr[i, 1].ToString() == "" || ComisAr[i, 1].ToString() == " " || ComisAr[i, 1].ToString() == null)
                    continue;
                
                bool first = false;
                bool second = false;
                bool third = false;
                bool from4to6 = false;
                bool from7to12 = false;
                bool abonent = false;
                bool regula = false;

                bool abonentAll = false;
                bool regulaAll = false;

                if (abonents.Contains(ComisAr[i, 4].ToString()))
                {
                    abonentAll = true;
                }
                else if (regular.Contains(ComisAr[i, 4].ToString()))
                {
                    regulaAll = true;
                }
                else
                    if (!NotFound.Contains(ComisAr[i, 4].ToString()))
                    NotFound += ComisAr[i, 4].ToString() + " ;   ";


                double nach = Convert.ToDouble(ComisAr[i, 3]);

                if (nach >= 120)
                {
                    if (abonents.Contains(ComisAr[i, 4].ToString()))
                    {
                        abonent = true;
                    }
                    else if (regular.Contains(ComisAr[i, 4].ToString()))
                    {
                        regula = true;
                    }
                    else
                        if (!NotFound.Contains(ComisAr[i, 4].ToString()))
                            NotFound += ComisAr[i, 4].ToString() + " ;   ";



                    switch (Convert.ToInt32(ComisAr[i, 2]))
                    {
                        case 1:
                            first = true;
                            break;
                        case 2:
                            second = true;
                            break;
                        case 3:
                            third = true;
                            break;
                        case 4:
                            from4to6 = true;
                            break;
                        case 5:
                            from4to6 = true;
                            break;
                        case 6:
                            from4to6 = true;
                            break;
                        case 7: case 8: case 9: case 10: case 11: case 12:
                            from7to12 = true;
                            break;
                    }
                }

                bool find = false;
                foreach (diler d in dilers)
                {
                    if (d.name.ToString() == ComisAr[i, 1].ToString())
                    {
                        d.sum += nach;

                        if (abonent)
                            d.Tab++;
                        else if (regula)
                            d.Treg++;

                        if (abonentAll)
                            d.TabAll++;
                        else if (regulaAll)
                            d.TregAll++;

                        find = true;

                        d.allincom++;
                        if (first)
                            d.count1201++;
                        if (second)
                            d.count1202++;
                        if (third)
                            d.count1203++; 
                        if (from4to6)
                            d.count12046++;
                        if (from7to12)
                            d.count120712++;
                        break;
                    }
                }

                if (!find)
                {
                    dilers.Add(new diler(ComisAr[i, 1],first,second,third,from4to6,from7to12,nach,abonent,regula,abonentAll,regulaAll));
                }
            }

            win.MessageBox.Show("1-ый этап завершен");



            string BasePath = new OpenExcelFile().Filenamereturn();
            if (BasePath == "can not open file")
                return;

            object[][] basearr = getbasearray(BasePath);
            int Nbase = basearr[0].Length;

            //win.MessageBox.Show(basearr[0][1].ToString());

            for (int i = 0; i < Nbase; i ++)
            {
                foreach (diler d in dilers)
                {
                    if (basearr[0][i].ToString() == d.name.ToString())
                        d.b += Convert.ToInt32(basearr[1][i]);
                }
            }

            win.MessageBox.Show("Конец 2.1-го этапа");


            BasePath = new OpenExcelFile().Filenamereturn();
            if (BasePath == "can not open file")
                return;

            basearr = getbasearray(BasePath);
            Nbase = basearr[0].Length;

            //win.MessageBox.Show(basearr[0][1].ToString());

            for (int i = 0; i < Nbase; i++)
            {
                foreach (diler d in dilers)
                {
                    if (basearr[0][i].ToString() == d.name.ToString())
                        d.a += Convert.ToInt32(basearr[1][i]);
                }
            }

            win.MessageBox.Show("Конец 2.2-го этапа");


            BasePath = new OpenExcelFile().Filenamereturn();
            if (BasePath == "can not open file")
                return;

            basearr = getbasearray(BasePath);
            Nbase = basearr[0].Length;

            //win.MessageBox.Show(basearr[0][1].ToString());

            for (int i = 0; i < Nbase; i++)
            {
                foreach (diler d in dilers)
                {
                    if (basearr[0][i].ToString() == d.name.ToString())
                        d.a += Convert.ToInt32(basearr[1][i]);
                }
            }

            win.MessageBox.Show("Конец 2.3-го этапа");


            BasePath = new OpenExcelFile().Filenamereturn();
            if (BasePath == "can not open file")
                return;

            basearr = getbasearray(BasePath);
            Nbase = basearr[0].Length;

            //win.MessageBox.Show(basearr[0][1].ToString());

            for (int i = 0; i < Nbase; i++)
            {
                foreach (diler d in dilers)
                {
                    if (basearr[0][i].ToString() == d.name.ToString())
                        d.a += Convert.ToInt32(basearr[1][i]);
                }
            }

            win.MessageBox.Show("Конец 2.4-го этапа");


            string toch = new OpenExcelFile().Filenamereturn();
            if (toch == "can not open file")
                return;
            object[][] tochki = getarray(toch, new int[] {3,4,5});


            object[,] result = new object[dilers.Count, 22];

            int k = 0;
            foreach (object t in tochki[0])
            {
                foreach (diler d in dilers)
                        if (t.ToString() == d.name.ToString())
                {
                    result[k, 0] = d.name;
                    result[k, 1] = d.b;
                    result[k, 2] = d.a;
                    result[k, 3] = d.b + d.a;
                    result[k, 4] = d.allincom;
                    result[k, 5] = d.sum;
                    result[k, 6] = d.count1201;
                    result[k, 7] = d.count1202;
                    result[k, 8] = d.count1203;
                    result[k, 9] = d.count12046;
                    result[k, 10] = d.count120712;
                    result[k, 11] = d.count1201 / Convert.ToDouble(d.allincom);
                    result[k, 12] = d.count1202 / Convert.ToDouble(d.allincom);
                    result[k, 13] = d.count1203 / Convert.ToDouble(d.allincom);
                    result[k, 14] = d.count12046 / Convert.ToDouble(d.allincom);
                    result[k, 15] = d.count120712 / Convert.ToDouble(d.allincom);
                    result[k, 16] = d.sum / Convert.ToDouble(d.allincom);
                    result[k, 17] = d.sum / Convert.ToDouble(d.a + d.b);
                    result[k, 18] = d.count1201 / Convert.ToDouble(d.a + d.b);
                    result[k, 19] = (d.count1201 + d.count1202 + d.count1203) / Convert.ToDouble(d.a + d.b);
                    result[k, 20] = (d.TabAll == 0)? 0: d.Tab/ Convert.ToDouble(d.TabAll);
                    result[k, 21] = (d.TregAll == 0)? 0:d.Treg/ Convert.ToDouble(d.TregAll);
                    k++;
                    break;
                }
            }


            string resPath = new OpenExcelFile().Filenamereturn();
            if (resPath == "can not open file")
                return;
            insert(resPath,result, dilers.Count,22);



            win.MessageBox.Show(NotFound,"Конец");
        }


        public void insert(string path,object[,] arr,int rows, int col)
        {

            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion

            range = null;
            range = excelworksheet.get_Range(R1C1[1] + "1:" + R1C1[col] + "1");
            range.Value2 = new object[,] { {"Дилер Дистр", "Кол-во в Базе" , "Кол-во в Архиве" , "Всего отгрузок" ,"Кол-во симок в комиссии", "Всего платежей" ,
                    "кол-во симок >120р в первом месяце" ,"кол-во симок >120р во втором месяце","кол-во симок >120р в третьем месяце",
             "кол-во симок >120р в 4-6 месяце"  ,"кол-во симок >120р в 7-12 месяце", "1) 1M" ,"2) 2M" ,"3) 3M" ,"4) 4-6M","5) 7-12M","6) платежи на комис" ,"7) платежи на отгрузки" ,
                    "8) хорошие (>120р) симки 1-го пер набл на кол-во отгрузок" ,"9) хорошие (>120р) симки 1,2,3 пер набл на кол-во отгрузок" } };

            range = null;
            range = excelworksheet.get_Range(R1C1[1] + "2:" + R1C1[col] + rows.ToString());
            range.Value2 = arr;

            #region Закрытие Excel

            book.Save();
            book.Close(false, false, false);

            ExcelApp.Quit();

            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            //workbooks = null;
            book = null;
            range = null;
            #endregion

        }


        private int dateinper(string s)
        {
            s = s.Remove(0,3);
            int month = Convert.ToInt32(s.Substring(0,2));
            int per = (s.Contains("2016")) ? 5 + (12 - month) : 5 - month;
            return per;
        }



        private void Button_Click_2(object sender, RoutedEventArgs e)
        {


            //win.MessageBox.Show(dateinper("06.01.2017 20:25").ToString());
            //return;


            string ComisPath = new OpenExcelFile().Filenamereturn();
            if (ComisPath == "can not open file")
                return;


            //object[][] ali = getarray(ComisPath, 23, 56, 15,10);
            object[][] ali = getarray(ComisPath, new int[] {27,95,98,19 });

            int N = ali[0].Length;
            object[,] ins = new object[N, 4];

            win.MessageBox.Show(N.ToString());


            for (int i = 0; i < N; i++)
            {
                ins[i, 0] = ali[1][i];
                ins[i, 2] = ali[0][i];
                ins[i, 1] = ali[2][i];
                ins[i, 3] = ali[3][i];
            }


            //for (int i = 0; i < N; i++ )    МЕГА
            //{
            //    ins[i, 0] = ali[2][i];
            //    ins[i, 1] = dateinper(ali[0][i].ToString()) - 1;
            //    ins[i, 2] = (ali[1][i] == null) ? 11.11: ali[1][i];
            //    ins[i, 3] = ali[3][i];
            //}

            string resPath = new OpenExcelFile().Filenamereturn();
            if (resPath == "can not open file")
                return;
            insert(resPath, ins, N, 4);

            win.MessageBox.Show(ali[0][10].ToString());

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {

           

            PdfFilePrinter.AdobeReaderPath = Textbox1.Text;
            

            string toch = new OpenExcelFile().Filenamereturn();
            if (toch == "can not open file")
                return;
            object[][] tochki = getarray(toch, new int[] { 1});

            string wpath = DirSelect();
            string[] files = Directory.GetFiles(wpath);
            int ad = 0;

            foreach (object t in tochki[0])
            {
                foreach(string file in files)
                {
                    if (file.Contains(t.ToString()))
                    {
                        try
                        {
                            PdfFilePrinter printer = new PdfFilePrinter(file, "HP LaserJet Professional P1606dn");
                            printer.Print();
                        }
                        catch (Exception ex)
                        {
                            win.MessageBox.Show("Error: " + ex.Message, "HP LaserJet Professional P1606dn");
                        }
                        ad++;
                        break;
                        // Process.Start(@"winword.exe", string.Format(@"{0} /mFilePrintDefault /mFileExit", wpath));
                    }
                }
                //if (ad > 3)
                //    break;
            }

            win.MessageBox.Show("Конец  " + ad.ToString());
        }
    }
}
