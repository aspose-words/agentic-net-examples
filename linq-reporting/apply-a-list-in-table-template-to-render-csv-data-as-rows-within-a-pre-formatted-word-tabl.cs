using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for .NET Core (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string csvPath = Path.Combine(outputDir, "Data.csv");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Id,Name,Age");
            writer.WriteLine("1,John Doe,30");
            writer.WriteLine("2,Jane Smith,25");
            writer.WriteLine("3,Bob Johnson,40");
        }

        // -----------------------------------------------------------------
        // 2. Build the template document with a pre‑formatted table and LINQ tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Optional title.
        builder.Writeln("Employee List");
        builder.Writeln();

        // Begin the foreach block that will repeat rows for each CSV record.
        builder.Writeln("<<foreach [row in data]>>");

        // Start the table.
        Table table = builder.StartTable();

        // Header row (static).
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row (repeated).
        builder.InsertCell();
        builder.Writeln("<<[row.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.HasHeaders = true; // Explicit for clarity.

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the data source named "data".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
