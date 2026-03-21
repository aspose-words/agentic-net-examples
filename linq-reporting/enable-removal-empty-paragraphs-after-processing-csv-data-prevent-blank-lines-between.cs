using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsCsvMailMerge
{
    class Program
    {
        static void Main()
        {
            // Use the current directory for all files so the example works anywhere.
            string baseDir = AppContext.BaseDirectory;
            string templatePath = Path.Combine(baseDir, "Template.docx");
            string csvPath = Path.Combine(baseDir, "Data.csv");
            string outputPath = Path.Combine(baseDir, "MergedResult.docx");

            // Ensure a simple template exists.
            if (!File.Exists(templatePath))
            {
                var templateDoc = new Document();
                var builder = new DocumentBuilder(templateDoc);
                builder.Writeln("Dear <<Name>>,");
                builder.Writeln("Your order <<OrderId>> has been shipped on <<ShipDate>>.");
                builder.Writeln("Thank you!");
                templateDoc.Save(templatePath);
            }

            // Ensure a simple CSV file exists.
            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath,
@"Name,OrderId,ShipDate
John Doe,12345,2023-01-15
Jane Smith,67890,2023-02-20");
            }

            // Load the Word template.
            Document doc = new Document(templatePath);

            // Load CSV data into a DataTable.
            DataTable csvTable = LoadCsvIntoDataTable(csvPath);

            // Remove empty paragraphs that may be created when a row contains empty fields.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

            // Execute mail merge. The column names in the CSV become merge field names.
            doc.MailMerge.Execute(csvTable);

            // Save the merged document.
            doc.Save(outputPath);

            Console.WriteLine($"Merged document saved to: {outputPath}");
        }

        /// <summary>
        /// Reads a CSV file and returns a DataTable where the first row is used as column headers.
        /// </summary>
        private static DataTable LoadCsvIntoDataTable(string csvPath)
        {
            var table = new DataTable();
            using (var reader = new StreamReader(csvPath))
            {
                bool isFirstLine = true;
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                        continue; // Skip empty lines and comment lines.

                    var fields = ParseCsvLine(line);
                    if (isFirstLine)
                    {
                        // Create columns from the header row.
                        foreach (var header in fields)
                            table.Columns.Add(header.Trim(), typeof(string));
                        isFirstLine = false;
                    }
                    else
                    {
                        // Add data rows.
                        var row = table.NewRow();
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            row[i] = i < fields.Length ? fields[i] : null;
                        }
                        table.Rows.Add(row);
                    }
                }
            }
            return table;
        }

        /// <summary>
        /// Very simple CSV parser that respects quoted fields and commas.
        /// For production code consider using a library such as CsvHelper.
        /// </summary>
        private static string[] ParseCsvLine(string line)
        {
            var result = new System.Collections.Generic.List<string>();
            bool inQuotes = false;
            var value = string.Empty;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    // Toggle quoting state. Two consecutive quotes inside a quoted field represent a literal quote.
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        value += '"';
                        i++; // Skip the escaped quote.
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(value);
                    value = string.Empty;
                }
                else
                {
                    value += c;
                }
            }
            result.Add(value);
            return result.ToArray();
        }
    }
}
