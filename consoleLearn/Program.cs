using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace consoleLearn
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Exsel faylinin adini yazin:");
            string Fname = Console.ReadLine();
            Console.WriteLine(Fname);

            Console.WriteLine("Fayil yaratmaq isteyirsinizse create yazin");
            string Create= Console.ReadLine();

            if (Create=="create")
            {
                Excel.Application xlApp = new
                Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                }
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                Console.WriteLine("Columlarin adlarini daxil edin (data elave etmek ucun stop yazin)");
                List<string> new_list = new List<string>();
                string colum;
                for (int i = 1; i < 10000; i++)
                {
                    colum = Console.ReadLine();
                    if (colum != "stop")
                    {
                        xlWorkSheet.Cells[1, i] = colum;
                        new_list.Add(colum);
                    }
                    else
                    {
                        break;
                    }
                }


                foreach (var item in new_list)
                {
                    Console.Write(item + ",");                    
                }
                Console.WriteLine("");
                Console.WriteLine("Yuxardaki ardiciliqla Datalari daxil edin");

                string besdir;
                for (int i = 2; i < 10000; i++)
                {                    
                    for (int x = 1; x < new_list.Count+1; x++)
                    {
                        xlWorkSheet.Cells[i, x] = Console.ReadLine();
                    }
                    Console.WriteLine("Yeni data elave etmek isteyirsiniz? Yes/no");
                    besdir = Console.ReadLine();
                    if (besdir == "no")
                    {
                        break;
                    }

                }
                

                xlWorkBook.SaveAs("d:\\"+ Fname + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            
        } 
    }
}
