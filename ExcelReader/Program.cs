using System;
using System.Drawing.Text;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;

namespace ExcelReadWrite
{
    class Program
    {
        static void Main(string[] args)
        {
            Boolean outLoop = true; 
            while (outLoop == true)
            {
                

                //need to loop through the items in the excel and find the next item to be worked on.
                string userInput = Console.ReadLine().ToLower();

                if (userInput == "a")
                {
                    ReturnAll();
                }
                if (userInput == "od")
                {
                    NextOnDeck();
                }
                if (userInput == "ct")
                {
                    ShowRecordsInClientTesting();
                }
                if (userInput == "ls")
                {
                    Console.WriteLine("a: Return All\nod: Next On Deck\nct: Show Records In Client Testing\nc: Close");
                }
                if(userInput == "c")
                {
                    outLoop = false;

                }

            }

            void ReturnAll()
            {
                using (var package = new ExcelPackage("C:/Users/Buddy/Desktop/test Data.xlsx"))
                {
                    var worksheet = package.Workbook.Worksheets["Sheet1"];

                    foreach (var task in worksheet.Cells["A2:A100"])
                    {
                        if (
                            worksheet.Cells[$"{task}"].Value != null
                            || worksheet.Cells[$"{task}"].Value.ToString() != ""
                        )
                        {
                            Console.WriteLine(worksheet.Cells[$"{task}"].Value.ToString());
                        }
                    }
                }
            }

            void NextOnDeck()
            {
                using (var package = new ExcelPackage("C:/Users/Buddy/Desktop/test Data.xlsx"))
                {
                    var worksheet = package.Workbook.Worksheets["Sheet1"];
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            if (
                                worksheet.Cells[row, col].Value != null
                                && !string.IsNullOrWhiteSpace(
                                    worksheet.Cells[row, col].Value.ToString()
                                )
                            )
                            {
                                Console.WriteLine($"{worksheet.Cells[row, col].Value}");
                                return;
                            }
                        }
                    }
                }
            }

            void ShowRecordsInClientTesting()
            {
                return;
            }
        }
    }
}
