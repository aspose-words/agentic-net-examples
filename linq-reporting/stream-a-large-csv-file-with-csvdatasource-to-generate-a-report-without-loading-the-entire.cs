using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the sample files.
        string csvPath = "data.csv";
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a large CSV file on disk (streamed later, not loaded fully).
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Write header.
            writer.WriteLine("Id,Name,Age");

            // Write many rows (e.g., 10,000 rows).
            for (int i = 1; i <= 10000; i++)
            {
                writer.WriteLine($"{i},Name{i},{20 + i % 30}");
            }
        }

        // -----------------------------------------------------------------
        // 2. Create a Word template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Id: <<[p.Id]>>, Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ','; // Default, but set explicitly.

        // Open the CSV file as a stream and create a CsvDataSource.
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, dataSource, "persons");
        }

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);

        // Optional: indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
