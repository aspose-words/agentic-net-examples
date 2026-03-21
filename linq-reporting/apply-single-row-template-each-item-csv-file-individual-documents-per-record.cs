using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

class CsvMailMergePerRecord
{
    static void Main()
    {
        // Use paths relative to the current directory so they work everywhere.
        string baseDir = AppContext.BaseDirectory;
        string templatePath = Path.Combine(baseDir, "SingleRowTemplate.docx");
        string csvPath = Path.Combine(baseDir, "Records.csv");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure the template exists – create a simple one if it does not.
        if (!File.Exists(templatePath))
        {
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Record:");
            builder.InsertField("MERGEFIELD Id");
            builder.Writeln();
            builder.InsertField("MERGEFIELD Name");
            templateDoc.Save(templatePath);
        }

        // Ensure the CSV file exists – create a small sample if it does not.
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath, "Id,Name\n1,John Doe\n2,Jane Smith\n3,Bob Johnson");
        }

        // Load the template document (single‑row mail‑merge template).
        Document template = new Document(templatePath);

        // Load CSV data into a DataTable.
        DataTable table = LoadCsvIntoDataTable(csvPath);

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Prepare field names for mail merge.
        string[] fieldNames = table.Columns.Cast<DataColumn>()
                                          .Select(col => col.ColumnName)
                                          .ToArray();

        // Iterate over each row and generate an individual document.
        foreach (DataRow row in table.Rows)
        {
            // Clone the template so each document starts from the same base.
            Document doc = (Document)template.Clone(true);

            // Perform mail merge for the current row.
            doc.MailMerge.Execute(fieldNames, row.ItemArray);

            // Build a file name using the first column value (assumed unique).
            string fileName = $"Document_{row[0]}.docx";
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the generated document.
            doc.Save(outputPath);
        }

        Console.WriteLine($"Generated {table.Rows.Count} documents in \"{outputFolder}\".");
    }

    // Reads a CSV file (with a header row) into a DataTable.
    private static DataTable LoadCsvIntoDataTable(string csvFilePath)
    {
        var dt = new DataTable();

        using (var reader = new StreamReader(csvFilePath))
        {
            bool hasHeaders = true;
            string[] headers = null;

            if (hasHeaders && !reader.EndOfStream)
            {
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
