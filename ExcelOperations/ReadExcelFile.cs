using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ExcelEditorApp.ExcelOperations
{
    class ReadExcelFile
    {
        public static void ReadFile ()
        {
            Console.Write("Enter the excel filename you want to view: ");
            string file = Console.ReadLine()!;
            string filePath = $"{file}.xlsx";

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                if (!sheetData.Elements<Row>().Any())
                {
                    Console.WriteLine("Current Excel file is empty!");
                }
                else
                {
                    Console.WriteLine("\nThis is the current content of your file");

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            Console.Write(cell.CellValue.Text + "\t");
                        }
                        Console.WriteLine();
                    }
                }
            }
        }
    }
}
