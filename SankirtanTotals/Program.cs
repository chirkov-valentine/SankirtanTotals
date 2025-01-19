using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab

namespace SankirtanTotals
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //CreateReport(new List<RowItem>());
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"E:\\Документы\\Распространение книг\\2024\\Totals.xlsx");
            var sheetsCount = xlWorkbook.Sheets.Count;
            
            List<string> allFioList = new List<string>();

            for (int i = 1; i <= sheetsCount; i++)
            {
                allFioList.AddRange(GetWorkSheetNames(xlWorkbook.Sheets[i]));
            }

            var fioList = allFioList.Distinct().ToList();

            List<RowItem> totalList = new List<RowItem>();
            foreach (var fio in fioList)
            {
                var totalItem = new RowItem() { FIO = fio };
                var rowSheet1 = FindRow(xlWorkbook.Sheets[1], fio);
                var rowSheet2 = FindRow(xlWorkbook.Sheets[2], fio);
                var rowSheet3 = FindRow(xlWorkbook.Sheets[3], fio);
                var rowSheet4 = FindRow(xlWorkbook.Sheets[4], fio);
                
                totalItem.H4 = rowSheet1.H4 
                               + rowSheet2.H4 
                               + rowSheet3.H4 
                               + rowSheet4.H4;
                totalItem.H3 = rowSheet1.H3 
                               + rowSheet2.H3 
                               + rowSheet3.H3 
                               + rowSheet4.H3;
                totalItem.S2 = rowSheet1.S2 
                               + rowSheet2.S2
                               + rowSheet3.S2 
                               + rowSheet4.S2;
                totalItem.S1 = rowSheet1.S1 
                               + rowSheet2.S1 
                               + rowSheet3.S1 
                               + rowSheet4.S1;
                totalItem.SBSets = rowSheet1.SBSets 
                                   + rowSheet2.SBSets
                                   + rowSheet3.SBSets
                                   + rowSheet4.SBSets;
                totalItem.CCSets = rowSheet1.CCSets 
                                   + rowSheet2.CCSets
                                   + rowSheet3.CCSets 
                                   + rowSheet4.CCSets;
                totalItem.Books = rowSheet1.Books 
                                  + rowSheet2.Books
                                  + rowSheet3.Books 
                                  + rowSheet4.Books;
                totalItem.Points = rowSheet1.Points 
                                   + rowSheet2.Points 
                                   + rowSheet3.Points 
                                   + rowSheet4.Points;

                Console.WriteLine($"{totalItem.FIO} {totalItem.H4} {totalItem.H3} {totalItem.S2} {totalItem.S1} {totalItem.SBSets} {totalItem.CCSets} {totalItem.Books} {totalItem.Points}");
                totalList.Add(totalItem);
            }

            totalList = totalList.OrderByDescending(t => t.Points).ToList();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            //Marshal.ReleaseComObject(xlRange5);
            //Marshal.ReleaseComObject(xlWorksheet5);

            for (int i = 1; i <= sheetsCount; i++)
            {
                Marshal.ReleaseComObject(xlWorkbook.Sheets[i]);
            }
            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            CreateReport(totalList);
        }

        private static IEnumerable<string> GetWorkSheetNames(Worksheet worksheet)
        {
            Excel.Range xlRange = worksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    yield return xlRange.Cells[i, 1].Value2.ToString();
            }
            Marshal.ReleaseComObject(xlRange);
        }

        private static RowItem FindRow(Excel._Worksheet worksheet, string fio)
        {
            Excel.Range xlRange = worksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int foundRow = -1;
            for (int i = 1; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    if (xlRange.Cells[i, 1].Value2.ToString() == fio)
                    { 
                      foundRow = i; 
                      break; 
                    }
               
            }
            RowItem result = new RowItem() { FIO = fio };
            if (foundRow != -1)
            {
                try
                {
                    result.H4 = Convert.ToInt32(xlRange.Cells[foundRow, 2]?.Value2?.ToString() ?? 0);
                }
                catch 
                {
                    result.H4 = 0;
                }
                try
                {
                    result.H3 = Convert.ToInt32(xlRange.Cells[foundRow, 3]?.Value2?.ToString() ?? 0);
                }
                catch 
                {
                    result.H3 = 0;
                }
                try
                {
                    result.S2 = Convert.ToInt32(xlRange.Cells[foundRow, 4]?.Value2.ToString() ?? 0);
                }
                catch
                {
                    result.S2 = 0;
                }
                try
                {
                    result.S1 = Convert.ToInt32(xlRange.Cells[foundRow, 5]?.Value2.ToString() ?? 0);
                }
                catch
                {
                    result.S1 = 0;
                }
                try
                {
                    result.SBSets = Convert.ToInt32(xlRange.Cells[foundRow, 7]?.Value2.ToString() ?? 0);
                } catch
                {
                    result.SBSets = 0;
                }
                try
                {
                    result.CCSets = Convert.ToInt32(xlRange.Cells[foundRow, 8]?.Value2?.ToString() ?? 0);
                }
                catch
                {
                    result.CCSets = 0;
                }
                try
                {
                    result.Books = Convert.ToInt32(xlRange.Cells[foundRow, 9]?.Value2?.ToString() ?? 0);
                }
                catch
                {
                    result.Books = 0;
                }
                try
                {
                    result.Points = float.Parse(xlRange.Cells[foundRow, 10]?.Value2.ToString() ?? 0);
                }
                catch 
                { 
                    result.Points = 0; 
                }
            }
            Marshal.ReleaseComObject(xlRange);
            return result;
        }

        static void CreateReport(List<RowItem> rows)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Name";
            xlWorkSheet.Cells[1, 2] = "H4";
            xlWorkSheet.Cells[1, 3] = "H3";
            xlWorkSheet.Cells[1, 4] = "S2";
            xlWorkSheet.Cells[1, 5] = "S1";
            xlWorkSheet.Cells[1, 6] = "SBSets";
            xlWorkSheet.Cells[1, 7] = "CCSets";
            xlWorkSheet.Cells[1, 8] = "Books";
            xlWorkSheet.Cells[1, 9] = "Points";

            int i = 2;
            foreach (RowItem row in rows) 
            {
                xlWorkSheet.Cells[i, 1] = row.FIO;
                xlWorkSheet.Cells[i, 2] = row.H4;
                xlWorkSheet.Cells[i, 3] = row.H3;
                xlWorkSheet.Cells[i, 4] = row.S2;
                xlWorkSheet.Cells[i, 5] = row.S1;
                xlWorkSheet.Cells[i, 6] = row.SBSets;
                xlWorkSheet.Cells[i, 7] = row.CCSets;
                xlWorkSheet.Cells[i, 8] = row.Books;
                xlWorkSheet.Cells[i, 9] = row.Points;
                i++;
            }



            xlWorkBook.SaveAs(@"E:\Документы\Распространение книг\2024\Итоги2023.xlsx", 
                Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, 
                Type.Missing, Type.Missing,
        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
