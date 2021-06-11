using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Aseo
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int j = 1;
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Aseo\\Pagos\\Excel4.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //for (rCnt = 2; rCnt <= 864; rCnt++)
                //{
                //    string cuota = (string)(range.Cells[rCnt, 6] as Excel.Range).Value2;
                //    int anno = Convert.ToInt32((range.Cells[rCnt, 7] as Excel.Range).Value2);
                //    if (Convert.ToInt32((range.Cells[rCnt, 4] as Excel.Range).Value2) != 2825)
                //    {
                //        string rol11 = "";
                //        string rol22 = "";
                //        int rol1 = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString().Length;
                //        int rol2 = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString().Length;
                //        if (rol1 < 5)
                //        {
                //            int aux = 5 - rol1;
                //            for (int l = 1; l <= aux; l++)
                //            {
                //                rol11 = rol11 + "0";
                //            }
                //        }
                //        if (rol2 < 5)
                //        {
                //            int aux = 5 - rol2;
                //            for (int l = 1; l <= aux; l++)
                //            {
                //                rol22 = rol22 + "0";
                //            }
                //        }
                //        if (cuota.IndexOf("1") != -1)
                //        {
                //            xlWorkSheet.Cells[j, 1] = rol11 + ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString() + "-" + rol22 + ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                //            xlWorkSheet.Cells[j, 2] = "30/04/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "1";
                //            j = j + 1;
                //        }
                //        if (cuota.IndexOf("2") != -1)
                //        {
                //            xlWorkSheet.Cells[j, 1] = rol11 + ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString() + "-" + rol22 + ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                //            xlWorkSheet.Cells[j, 2] = "30/06/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "2";
                //            j = j + 1;
                //        }
                //        if (cuota.IndexOf("3") != -1)
                //        {
                //            xlWorkSheet.Cells[j, 1] = rol11 + ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString() + "-" + rol22 + ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                //            xlWorkSheet.Cells[j, 2] = "30/09/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "3";
                //            j = j + 1;
                //        }
                //        if (cuota.IndexOf("4") != -1)
                //        {
                //            xlWorkSheet.Cells[j, 1] = rol11 + ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString() + "-" + rol22 + ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                //            xlWorkSheet.Cells[j, 2] = "30/11/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "4";
                //            j = j + 1;
                //        }
                //    }
                //    else
                //    {
                //        for (int k = 3; k <= 238; k++)
                //        {
                //            string aux = "";
                //            if (k < 100)
                //            {
                //                if (k < 10) aux = "0000" + k;
                //                else aux = "000" + k;
                //            }
                //            else
                //            {
                //                aux = "00" + k;
                //            }

                //            xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                //            xlWorkSheet.Cells[j, 2] = "30/04/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "1";
                //            j = j + 1;
                //            xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                //            xlWorkSheet.Cells[j, 2] = "30/06/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "2";
                //            j = j + 1;
                //            xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                //            xlWorkSheet.Cells[j, 2] = "30/09/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "3";
                //            j = j + 1;
                //            xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                //            xlWorkSheet.Cells[j, 2] = "30/11/" + anno;
                //            xlWorkSheet.Cells[j, 3] = "4";
                //            j = j + 1;
                //        }
                //    }
                //}

                for (rCnt = 2; rCnt <= 183; rCnt++)
                {
                    string rol = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                    string cuota = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    int anno = Convert.ToInt32((range.Cells[rCnt, 5] as Excel.Range).Value2);

                    //if (Convert.ToInt32((range.Cells[rCnt, 4] as Excel.Range).Value2) != 2825)
                    //{
                    string rol1 = rol.Substring(0, rol.IndexOf("-") + 1);
                    string rol2 = rol.Substring(rol.IndexOf("-") + 1);
                    if (rol1.Length < 6)
                    {
                        int aux = 6 - rol1.Length;
                        for (int k = 0; k < aux; k++)
                        {
                            rol1 = "0" + rol1;
                        }
                    }
                    if (rol2.Length < 5)
                    {
                        int aux = 5 - rol2.Length;
                        for (int k = 0; k < aux; k++)
                        {
                            rol2 = "0" + rol2;
                        }
                    }

                    if (cuota.IndexOf("1") != -1)
                    {
                        xlWorkSheet.Cells[j, 1] = rol1 + rol2;
                        xlWorkSheet.Cells[j, 2] = "30/04/" + anno;
                        xlWorkSheet.Cells[j, 3] = "1";
                        j = j + 1;
                    }
                    if (cuota.IndexOf("2") != -1)
                    {
                        xlWorkSheet.Cells[j, 1] = rol1 + rol2;
                        xlWorkSheet.Cells[j, 2] = "30/06/" + anno;
                        xlWorkSheet.Cells[j, 3] = "2";
                        j = j + 1;
                    }
                    if (cuota.IndexOf("3") != -1)
                    {
                        xlWorkSheet.Cells[j, 1] = rol1 + rol2;
                        xlWorkSheet.Cells[j, 2] = "30/09/" + anno;
                        xlWorkSheet.Cells[j, 3] = "3";
                        j = j + 1;
                    }
                    if (cuota.IndexOf("4") != -1)
                    {
                        xlWorkSheet.Cells[j, 1] = rol1 + rol2;
                        xlWorkSheet.Cells[j, 2] = "30/11/" + anno;
                        xlWorkSheet.Cells[j, 3] = "4";
                        j = j + 1;
                    }
                    //}
                    //else
                    //{
                    //    for (int k = 3; k <= 238; k++)
                    //    {
                    //        string aux = "";
                    //        if (k < 100)
                    //        {
                    //            if (k < 10) aux = "0000" + k;
                    //            else aux = "000" + k;
                    //        }
                    //        else
                    //        {
                    //            aux = "00" + k;
                    //        }

                    //        xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                    //        xlWorkSheet.Cells[j, 2] = "30/04/" + anno;
                    //        xlWorkSheet.Cells[j, 3] = "1";
                    //        j = j + 1;
                    //        xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                    //        xlWorkSheet.Cells[j, 2] = "30/06/" + anno;
                    //        xlWorkSheet.Cells[j, 3] = "2";
                    //        j = j + 1;
                    //        xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                    //        xlWorkSheet.Cells[j, 2] = "30/09/" + anno;
                    //        xlWorkSheet.Cells[j, 3] = "3";
                    //        j = j + 1;
                    //        xlWorkSheet.Cells[j, 1] = "02825-" + aux;
                    //        xlWorkSheet.Cells[j, 2] = "30/11/" + anno;
                    //        xlWorkSheet.Cells[j, 3] = "4";
                    //        j = j + 1;
                    //    }
                    //}
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);






            xlWorkBook.SaveAs("D:\\Aseo\\Pagos\\Finalaseo.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
