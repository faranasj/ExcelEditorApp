using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ExcelEditorApp.ExcelOperations
{
    class ReadExcelFile
    {
        public static void ReadFile ()
        {
            Console.WriteLine("Enter the excel filename you want to view: ");
            string filePath = Console.ReadLine()!;

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
                    Console.WriteLine("This is the current content of your file");

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
