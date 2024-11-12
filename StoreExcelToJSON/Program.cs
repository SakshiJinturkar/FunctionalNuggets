using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Newtonsoft.Json;
using System.Linq;
using System.Text.RegularExpressions;

class Program
{
    static void Main(string[] args)
    {
        // Set the LicenseContext
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or .Commercial if you have a commercial license

        string excelFilePath = @""; // Replace with your Excel file path
        string jsonFolderPath = @""; // Replace with your JSON folder path

        List<string> keysToExtract = new List<string>
        {
            "Article number",
            "Article name",
            "Manufacturer",
            "Manufacturer number",
            "Package size",
            "Measurement unit",
            "Price unit",
            "Measurement type",
            "Booking account",
            "Base price",
            "Discount group",
            "Weight",
            "Length",
            "Width",
            "Height",
            "Thickness",
            "Circumferential processing",
            "Mechanical processing",
            "Package unit",
            "Price date",
            "Standard",
            "Anodic processing",
            "Color",
            "Material",
            "System",
            "Part list relevant",
            "Calculation relevant",
            "Make or buy",
            "Pattern",
            "Profile",
            "Calculation factor",
            "Storage type",
            "Tax",
            "Dynamic geometry",
            "Single part drawing"
        };

        // Load the Excel file
        FileInfo excelFile = new FileInfo(excelFilePath);
        using (ExcelPackage package = new ExcelPackage(excelFile))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

            int rows = worksheet.Dimension.Rows;
            int cols = worksheet.Dimension.Columns;

            int articleNumberColumn = 1; // Assuming article number is in the first column (A)

            Dictionary<string, Dictionary<string, string>> articleData = new Dictionary<string, Dictionary<string, string>>();

            // Read data from Excel and store it in a dictionary
            for (int row = 2; row <= rows; row++) // Assuming header is in the first row
            {
                string currentArticleNumber = worksheet.Cells[row, articleNumberColumn].Value?.ToString();

                if (!string.IsNullOrEmpty(currentArticleNumber) && !articleData.ContainsKey(currentArticleNumber))
                {
                    articleData[currentArticleNumber] = new Dictionary<string, string>();

                    for (int col = 1; col <= cols; col++)
                    {
                        string key = worksheet.Cells[1, col].Value?.ToString();
                        string value = worksheet.Cells[row, col].Value?.ToString() ?? "";

                        if (keysToExtract.Contains(key))
                        {
                            if (string.IsNullOrEmpty(value))
                            {
                                articleData[currentArticleNumber][key] = "\"\"";
                            }
                            else if (double.TryParse(value, out double numericValue))
                            {
                                articleData[currentArticleNumber][key] = numericValue.ToString(); // Numeric values without quotes
                            }
                            else if (DateTime.TryParse(value, out DateTime dateValue))
                            {
                                string dateVal = dateValue.ToString("dd/MM/yyyy");
                                dateVal = dateVal.Replace("-", "/");
                                articleData[currentArticleNumber][key] = dateVal; // Format date as MM/dd/yyyy
                            }
                            else
                            {
                                articleData[currentArticleNumber][key] = $"\"{value}\""; // String values with quotes
                            }
                        }
                    }
                }
            }

            // Ask user for input of article numbers line by line
            Console.WriteLine("Enter article numbers (one per line). Press Enter after each number. Press Enter twice to finish:");
            Console.ReadLine();
            List<string> articleNumbers = new List<string>();
            string articleNumberInput;

            while (!string.IsNullOrWhiteSpace(articleNumberInput = Console.ReadLine()))
            {
                articleNumbers.Add(articleNumberInput.Trim());
            }

            // Process article numbers and generate JSON for specified article numbers
            foreach (var articleNumber in articleNumbers)
            {
                if (articleData.ContainsKey(articleNumber))
                {
                    Dictionary<string, string> articleValues = articleData[articleNumber];

                    string json = GenerateJSON(articleValues, keysToExtract);
                    string jsonFilePath = Path.Combine(jsonFolderPath, $"{articleNumber}.json");

                    // Save JSON data to file
                    File.WriteAllText(jsonFilePath, json);
                    Console.WriteLine($"JSON file created for article number '{articleNumber}'");
                }
                else
                {
                    Console.WriteLine($"Article number '{articleNumber}' not found in the Excel file.");
                    Console.ReadLine();
                }
            }
        }
    }

    static string GenerateJSON(Dictionary<string, string> data, List<string> keys)
    {
        // Ensure JSON maintains the sequence of keys
        List<string> orderedKeys = new List<string>(keys);

        // Add default key-value pairs for missing keys
        var defaultKeyValues = new Dictionary<string, string>
    {
        { "Part list relevant", "true" },
        { "Calculation relevant", "true" },
        { "Make or buy", "\"buy\"" },
        { "Pattern", "0" },
        { "Profile", "\"\"" },
        { "Dynamic geometry", "false" },
        { "Single part drawing", "false" }
    };

        foreach (var pair in defaultKeyValues)
        {
            if (!data.ContainsKey(pair.Key))
            {
                data.Add(pair.Key, pair.Value);
                orderedKeys.Add(pair.Key);
            }
        }

        // Convert dictionary to JSON string maintaining sequence
        var orderedData = new Dictionary<string, string>();
        foreach (string key in orderedKeys)
        {
            orderedData[key] = data[key];
        }

        string indent = "    "; // Set the desired indentation (four spaces in this case)
        string[] keyValuePairs = orderedData.Select(kv => $"{indent}\"{kv.Key}\": {kv.Value}").ToArray();

        // Join key-value pairs with newline character and add braces for JSON formatting
        return "{\n" + string.Join(",\n", keyValuePairs) + "\n}";
    }
}
