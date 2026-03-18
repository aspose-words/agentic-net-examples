using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

class CsvMailMergePerRecord
{
    static void Main()
    {
        // Paths – adjust as needed.
        string templatePath = @"C:\Templates\SingleRowTemplate.docx";
        string csvPath = @"C:\Data\Records.csv";
        string outputFolder = @"C:\Output\";

        // Load the template document (single‑row mail‑merge template).
        Document template = new Document(templatePath);

        // Load CSV data into a DataTable.
        DataTable table = LoadCsvIntoDataTable(csvPath);

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Iterate over each row and generate an individual document.
        foreach (DataRow row in table.Rows)
        {
            // Clone the template so each document starts from the same base.
            Document doc = (Document)template.Clone(true);

            // Perform mail merge for the current row.
            doc.MailMerge.Execute(row);

            // Build a file name using a field value (e.g., "Id" column) or row index.
            string fileName = $"Document_{row[0]}.docx"; // assumes first column is unique identifier
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the generated document.
            doc.Save(outputPath);
        }
    }

    // Reads a CSV file (with optional header row) into a DataTable.
    private static DataTable LoadCsvIntoDataTable(string csvFilePath)
    {
        DataTable dt = new DataTable();

        using (var reader = new StreamReader(csvFilePath))
        {
            bool hasHeaders = true; // set to false if your CSV has no header row
            string[] headers = null;

            if (hasHeaders && !reader.EndOfStream)
            {
                // First line contains column names.
                headers = ParseCsvLine(reader.ReadLine());
                foreach (var header in headers)
                    dt.Columns.Add(header);
            }

            while (!reader.EndOfStream)
            {
                string[] fields = ParseCsvLine(reader.ReadLine());

                // If the CSV has no header row, create columns on the first data line.
                if (!hasHeaders && dt.Columns.Count == 0)
                {
                    for (int i = 0; i < fields.Length; i++)
                        dt.Columns.Add($"Column{i + 1}");
                }

                // Add the row to the table.
                dt.Rows.Add(fields);
            }
        }

        return dt;
    }

    // Simple CSV line parser handling commas and quoted fields.
    private static string[] ParseCsvLine(string line)
    {
        var fields = new System.Collections.Generic.List<string>();
        bool inQuotes = false;
        var field = new System.Text.StringBuilder();

        foreach (char c in line)
        {
            if (c == '\"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (c == ',' && !inQuotes)
            {
                fields.Add(field.ToString());
                field.Clear();
            }
            else
            {
                field.Append(c);
            }
        }

        fields.Add(field.ToString());
        return fields.ToArray();
    }
}
