using ExcelEditorApp.ExcelOperations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorApp.Menu
{
    class ExcelEditorMenu
    {
        public static void EditorMenu()
        {
            Console.WriteLine("Welcome to the Excel File Editor App!");

            var flag = true;

            while (flag)
            {
                PrintMenu();
                Console.Write("\nPlease enter your option: ");
                string option = Console.ReadLine()!;

                switch (option)
                {
                    case "1":
                        CreateExcelFile.CreateNewFile(); 
                        break;
                    case "2":
                        ReadExcelFile.ReadFile(); 
                        break;
                    case "3":
                        WriteToExcelFile.WriteToFile();
                        break;
                    case "0":
                        flag = false;
                        break;
                    default:
                        Console.WriteLine("Invalid input!");
                        break;
                }
            }
        }
        static void PrintMenu()
        {
            Console.WriteLine("Enter 1 to create a new excel file");
            Console.WriteLine("Enter 2 to read an excel file");
            Console.WriteLine("Enter 3 to write to an excel file");
            Console.WriteLine("Enter 0 to exit.");
        }
    }
}
