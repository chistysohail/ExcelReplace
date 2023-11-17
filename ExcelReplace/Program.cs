using ClosedXML.Excel;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelReplace
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load the Excel workbook
                using (var workbook = new XLWorkbook("path/to/your/excelfile.xlsx"))
                {
                    var worksheet = workbook.Worksheet(1); // assuming data is in the first worksheet

                    // Read the TypeScript file
                    string tsFilePath = "path/to/your/app.const.ts";
                    string fileContent = File.ReadAllText(tsFilePath);

                    // Iterate through the rows of the Excel file
                    foreach (var row in worksheet.RangeUsed().Rows())
                    {
                        string findValue = row.Cell(1).GetValue<string>();    // Column A
                        string replaceValue = row.Cell(2).GetValue<string>(); // Column B

                        Console.WriteLine($"Looking for: {findValue}, to replace with: {replaceValue}");

                        // Check if the word exists in the file
                        if (fileContent.Contains(findValue))
                        {
                            Console.WriteLine($"Found '{findValue}' in the file. Replacing...");
                            fileContent = Regex.Replace(fileContent, Regex.Escape(findValue), replaceValue);
                        }
                        else
                        {
                            Console.WriteLine($"'{findValue}' not found in the file.");
                        }
                    }

                    // Save the modified TypeScript file
                    File.WriteAllText(tsFilePath, fileContent);
                    Console.WriteLine("File saved successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
