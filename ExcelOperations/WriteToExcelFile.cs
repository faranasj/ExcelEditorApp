using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorApp.ExcelOperations
{
    class WriteToExcelFile
    {
        public static void WriteToFile()
        {
            Console.Write("Enter the excel filename you want to write to: ");
            string file = Console.ReadLine()!;
            string filePath = $"{file}.xlsx";

            Console.WriteLine("\nWould you like to add a new row? (y/n)");
            string response = Console.ReadLine()!;

            if (response.ToLower() == "y")
            {
                Console.WriteLine("Enter data separated by commas (e.g., 1,Jane,Doe): ");
                string[] rowData = Console.ReadLine()!.Split(',');

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Row newRow = new Row();

                    foreach (string cellData in rowData)
                    {
                        Cell cell = new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(cellData)
                        };
                        newRow.Append(cell);
                    }
                    sheetData.Append(newRow);
                }
                Console.WriteLine("");
                Console.WriteLine("Data added successfully...");
                Console.WriteLine("");
            }
        }
    }
}
