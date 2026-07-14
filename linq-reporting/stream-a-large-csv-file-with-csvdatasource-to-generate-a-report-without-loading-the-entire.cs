using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some CSV encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Path for a temporary large CSV file
        string csvPath = Path.Combine(Path.GetTempPath(), "large.csv");

        // Generate a CSV file with many rows (e.g., 5000 rows)
        GenerateCsv(csvPath, 5000);

        // Create a Word template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report generated from CSV data:");
        builder.Writeln("Id\tName\tAge");
        // LINQ Reporting foreach tag that iterates over the CSV rows
        builder.Writeln("<<foreach [row in persons]>>");
        builder.Writeln("<<[row.Id]>>\t<<[row.Name]>>\t<<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using a CsvDataSource that reads from a stream
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            // Load options: first line contains headers, comma delimiter
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            loadOptions.Delimiter = ',';

            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "persons");
        }

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportFromCsv.docx");
        template.Save(outputPath);

        // Clean up the temporary CSV file
        File.Delete(csvPath);
    }

    // Helper method to create a CSV file with the specified number of rows
    private static void GenerateCsv(string path, int rowCount)
    {
        using (StreamWriter writer = new StreamWriter(path))
        {
            writer.WriteLine("Id,Name,Age"); // Header
            for (int i = 1; i <= rowCount; i++)
            {
                writer.WriteLine($"{i},Name_{i},{20 + (i % 30)}");
            }
        }
    }
}
