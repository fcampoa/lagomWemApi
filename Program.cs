using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string path = "Copia de Libro1.xlsx";
                object oMissing = System.Reflection.Missing.Value;
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wb = excel.Workbooks.Open(path, oMissing, false);
                Excel.Worksheet worksheet = wb.Sheets["hoja1"];
                var value = ((Excel.Range)worksheet.Cells[1, "A"]).Value; ;
                Console.WriteLine(value);
                //excel.Run("Division");
                //excel.Calculate();
                wb.Save();
                wb.Close();
                excel.Quit();
                // excel.Visible = true;
            }catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
