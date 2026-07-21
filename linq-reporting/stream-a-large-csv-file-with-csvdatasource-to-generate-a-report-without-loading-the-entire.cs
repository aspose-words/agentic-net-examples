using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "Data.csv");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a large CSV file with headers "Name,Age" and many rows.
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Name,Age"); // Header row
            for (int i = 1; i <= 1000; i++)
            {
                writer.WriteLine($"Person {i},{20 + (i % 30)}");
            }
        }

        // ---------------------------------------------------------------
        // 2. Build a simple Word template containing LINQ Reporting tags.
        // ---------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from streamed CSV data:");
        // Begin a foreach block that iterates over the data source named "persons".
        builder.Writeln("<<foreach [row in persons]>>");
        // Output each row's fields.
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template back for report generation.
        // ---------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // ---------------------------------------------------------------
        // 4. Configure CSV data loading options (headers present, comma delimiter).
        // ---------------------------------------------------------------
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true)
        {
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };

        // ---------------------------------------------------------------
        // 5. Create a CsvDataSource that reads from the CSV file via a stream.
        // ---------------------------------------------------------------
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            var csvDataSource = new CsvDataSource(csvStream, loadOptions);

            // -----------------------------------------------------------
            // 6. Build the report using ReportingEngine.
            // -----------------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default behavior

            // The data source name used in the template tags is "persons".
            engine.BuildReport(reportDoc, csvDataSource, "persons");
        }

        // ---------------------------------------------------------------
        // 7. Save the generated report.
        // ---------------------------------------------------------------
        reportDoc.Save(reportPath);

        // Optional: inform the user (no interactive input required).
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
