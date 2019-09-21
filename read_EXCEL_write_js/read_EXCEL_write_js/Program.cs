using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace read_EXCEL_write_js
{
    class Program
    {
        static void Main(string[] args)
        {
            string createText = getExcel();
            File.WriteAllText(@"D:\GitSpace\Songs_List\test.js", createText);
            Console.ReadKey();
        }

        public static string getExcel ()
        {/*https://coderwall.com/p/app3ya/read-excel-file-in-c */
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\GitSpace\Songs_List\excel\song_list.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string result = "var list = [{" + Environment.NewLine; 
            string[] song = new string[4] { "\'id\': ", "\'name\': ", "\'author\': ", "\'words_count\': " };
            Console.WriteLine("執行中...., 請等ok出現");
            //Console.WriteLine("一列"+colCount+"個。");
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1 && i > 1) { 
                        //Console.Write("\r\n");
                        result += "},{" + Environment.NewLine;
                    }

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        if(j == 1)
                        {
                            result += "\t" + song[j - 1] + xlRange.Cells[i, j].Value2.ToString() + "," + Environment.NewLine;
                        }else if (j > 1 && j < colCount)
                        {
                            result += "\t" + song[j - 1] + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"," + Environment.NewLine;
                        }
                        else{
                            result += "\t" + song[j - 1] + xlRange.Cells[i, j].Value2.ToString() + Environment.NewLine;                            
                        }
                        
                        //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                }
            }
            result += "}];";
            Console.WriteLine("ok,隨便按顆鍵離開");
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return result;
        }
    }
}
