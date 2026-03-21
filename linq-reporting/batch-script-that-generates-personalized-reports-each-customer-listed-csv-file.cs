using System;
using System.Data;
using System.IO;

class BatchReportGenerator
{
    // Simple text template with placeholders that will be replaced by CSV column values.
    private const string TemplateContent = 
@"Customer Report
Name: {Name}
Email: {Email}
Customer ID: {CustomerID}
";

    // Path to the CSV file that contains customer data. First line must contain column headers.
    private const string CsvPath = @"Customers.csv";

    // Folder where the generated reports will be saved.
    private const string OutputFolder = @"Generated";

    static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(OutputFolder);

        // Ensure a sample CSV file exists.
        EnsureSampleCsvExists(CsvPath);

        // Load CSV data into a DataTable.
        DataTable customersTable = LoadCsvIntoDataTable(CsvPath);

        // For each customer row generate a separate report.
        foreach (DataRow row in customersTable.Rows)
        {
            // Build a file name using a unique identifier, e.g., CustomerID or Name.
            string fileName = $"{row["Name"]}_{row["CustomerID"]}.txt";
            string outputPath = Path.Combine(OutputFolder, fileName);

            // Create report content by replacing placeholders with actual values.
            string reportContent = TemplateContent;
            foreach (DataColumn column in customersTable.Columns)
            {
                string placeholder = $"{{{column.ColumnName}}}";
                reportContent = reportContent.Replace(placeholder, row[column].ToString());
            }

            // Save the generated report.
            File.WriteAllText(outputPath, reportContent);
        }

        Console.WriteLine("All reports have been generated.");
    }

    // Reads a CSV file and returns a DataTable.
    // Assumes the first line contains column headers.
    private static DataTable LoadCsvIntoDataTable(string csvFilePath)
    {
        DataTable table = new DataTable();

        using (var reader = new StreamReader(csvFilePath))
        {
            bool isFirstLine = true;
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] fields = line.Split(',');

                if (isFirstLine)
                {
                    // Create columns based on header names.
                    foreach (string header in fields)
                    {
                        table.Columns.Add(header.Trim());
                    }
                    isFirstLine = false;
                }
                else
                {
                    // Add a new row with the field values.
                    DataRow dataRow = table.NewRow();
                    for (int i = 0; i < fields.Length; i++)
                    {
                        dataRow[i] = fields[i].Trim();
                    }
                    table.Rows.Add(dataRow);
                }
            }
        }

        return table;
    }

    // Creates a simple CSV file with sample data if it does not already exist.
    private static void EnsureSampleCsvExists(string path)
    {
        if (File.Exists(path))
            return;

        var sampleData = new[]
        {
            "CustomerID,Name,Email",
            "1,John Doe,john.doe@example.com",
            "2,Jane Smith,jane.smith@example.com",
            "3,Bob Johnson,bob.johnson@example.com"
        };

        File.WriteAllLines(path, sampleData);
    }
}
