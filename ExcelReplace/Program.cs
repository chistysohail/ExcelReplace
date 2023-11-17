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

                    // Replace in the TypeScript file content
                    fileContent = Regex.Replace(fileContent, findValue, replaceValue);
                }

                // Save the modified TypeScript file
                File.WriteAllText(tsFilePath, fileContent);
            }
        }
    }
}
