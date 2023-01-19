using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace Test_Coolshop
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("Enter Path Index(start from 1) SearchName");
                string fullPath = Console.ReadLine();
                try
                {
                    List<dynamic> lstPath = new List<dynamic>(fullPath.Split(" "));
                    string path = lstPath[0];
                    int indexSearch = Int32.Parse(lstPath[1]);
                    string search = lstPath[2].ToString();
                    //Create COM Objects.
                    Application excelApp = new Application();


                    if (excelApp == null)
                    {
                        Console.WriteLine("Excel is not installed!!");
                        return;
                    }

                    Workbook excelBook = excelApp.Workbooks.Open(@path);
                    _Worksheet excelSheet = excelBook.Sheets[1];
                    Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                    int rows = excelRange.Rows.Count;
                    for (int i = 1; i <= rows; i++)
                    {
                        if (excelRange.Cells[i, indexSearch] != null)
                        {
                            //convert row value to complex
                            string rowValue = indexSearch == 4 ? DateTime.FromOADate(excelRange.Cells[i, indexSearch].Value2).ToString("MM/dd/yyyy") : excelRange.Cells[i, indexSearch].Value2.ToString();

                            //write the console                 
                            if (rowValue == search)
                            {
                                Console.Write(excelRange.Cells[i, 1].Value2 + "," +
                                        excelRange.Cells[i, 2].Value2 + "," +
                                        excelRange.Cells[i, 3].Value2 + "," +
                                       DateTime.FromOADate(excelRange.Cells[i, 4].Value2).ToString("MM/dd/yyyy") + ";");
                            }
                        }
                        else
                        {
                            Console.WriteLine("No Data");
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Error Inputs");
                }

                Console.ReadLine();
            }
        }
    }
}
