using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class BatchReportGenerator
{
    // Path to the template Word document that contains merge fields like <<[Customer.Name]>>, <<[Customer.Email]>> etc.
    private const string TemplatePath = @"C:\Reports\Template.docx";

    // Path to the CSV file that contains customer data. First line must contain column headers.
    private const string CsvPath = @"C:\Reports\Customers.csv";

    // Folder where the generated reports will be saved.
    private const string OutputFolder = @"C:\Reports\Generated";

    static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(OutputFolder);

        // Load the template document.
        Document template = new Document(TemplatePath);

        // Load CSV data into a DataTable.
        DataTable customersTable = LoadCsvIntoDataTable(CsvPath);

        // For each customer row generate a separate report.
        foreach (DataRow row in customersTable.Rows)
        {
            // Clone the template so each report starts from a fresh copy.
            Document report = (Document)template.Clone(true);

            // Perform mail merge using the column names as field names.
            // The field names in the template must match the column names in the CSV.
            report.MailMerge.Execute(row);

            // Build a file name using a unique identifier, e.g., CustomerID or Name.
            string fileName = $"{row["Name"]}_{row["CustomerID"]}.docx";
            string outputPath = Path.Combine(OutputFolder, fileName);

            // Save the generated report.
            report.Save(outputPath);
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
}
