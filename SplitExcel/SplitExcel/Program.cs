using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace SplitExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string FileName = string.Join("", args);        
            SplitAndMail(FileName);
        }

        static void SplitAndMail(string FileName)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;

            FileInfo fi = new FileInfo(FileName);
            string FullFileName = fi.FullName.ToString();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(FullFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            }
            catch
            {
                Console.WriteLine("You must choose a file");
                return;
            }
            xlWorkBook = xlApp.Workbooks.Open(FullFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Range rng = xlApp.get_Range("A1");
            //int index = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet displayWorksheet in xlApp.Worksheets)
            {
                
                string root = fi.Directory.ToString();
                string sheetName = displayWorksheet.Name.ToString();
                string SaveFileName = root + "\\" + sheetName + ".xlsx";
                //Microsoft.Office.Interop.Excel.Application NewxlApp;
                //NewxlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook NewWorkbook;
                NewWorkbook = xlApp.Workbooks.Add();

                Microsoft.Office.Interop.Excel.Worksheet NewWorkSheet;
                NewWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)NewWorkbook.Worksheets.Add();

                displayWorksheet.Copy(NewWorkSheet);
                NewWorkbook.SaveAs(SaveFileName);
                NewWorkbook.Save();
                NewWorkbook.Close();
            }
        }
    }
}
