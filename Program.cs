using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ExcelEditorApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the Excel File Editor App!");
            Console.WriteLine("Please provide the filename of your excel file (e.g. your-file.xlsx).");
            Console.WriteLine("Disclaimer - If no filename is given, 'your-file.xlsx' is used. Additionally, if file does not exist, a new one is created.");
            string filePath = Console.ReadLine()!;

            if (filePath == string.Empty)
            {
                filePath = "your-file.xlsx";
            }

            if (!File.Exists(filePath))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = doc.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                    sheets.Append(sheet);
                    workbookPart.Workbook.Save();
                }
            }
        }
    }
}