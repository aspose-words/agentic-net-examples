using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with a header row and a few records.
        // -----------------------------------------------------------------
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Age",
            "1,John Doe,30",
            "2,Jane Smith,25",
            "3,Bob Johnson,40"
        });

        // ---------------------------------------------------------------
        // 2. Build a Word template that uses LINQ Reporting tags.
        // ---------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("=== Detailed Record (ElementAt) ===");
        // The collection name "persons" will be supplied to the engine.
        // ElementAt(1) fetches the second record (zero‑based index).
        builder.Writeln("Id:   <<[persons.ElementAt(1).Id]>>");
        builder.Writeln("Name: <<[persons.ElementAt(1).Name]>>");
        builder.Writeln("Age:  <<[persons.ElementAt(1).Age]>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // ---------------------------------------------------------------
        var template = new Document(templatePath);

        // Configure CSV loading to treat the first line as headers.
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report. The root object name must match the tag prefix.
        var engine = new ReportingEngine();
        engine.BuildReport(template, csvDataSource, "persons");

        // ---------------------------------------------------------------
        // 4. Save the generated report.
        // ---------------------------------------------------------------
        template.Save(reportPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
