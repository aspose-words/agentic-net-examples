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

        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workDir, "people.csv");
        string templatePath = Path.Combine(workDir, "template.docx");
        string reportPath = Path.Combine(workDir, "report.docx");

        // 1. Create a CSV file with quoted fields that contain commas.
        // Header: Name,Address,Notes
        // Sample row: "John Doe","123 Main St, Apt 4","He said, ""Hello, world!"""
        string[] csvLines =
        {
            "Name,Address,Notes",
            "\"John Doe\",\"123 Main St, Apt 4\",\"He said, \"\"Hello, world!\"\"\"",
            "\"Jane Smith\",\"456 Oak Rd, Suite 5\",\"Note with, comma\""
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // 2. Create a Word template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin a foreach loop over the CSV rows (named "persons").
        builder.Writeln("<<foreach [row in persons]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Address: <<[row.Address]>>");
        builder.Writeln("Notes: <<[row.Notes]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template back (demonstrating the load step).
        Document loadedTemplate = new Document(templatePath);

        // 4. Configure CSV data load options.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // true => first line has headers
        {
            Delimiter = ',',          // default, but set explicitly
            QuoteChar = '"',          // default, but set explicitly
            CommentChar = '#',        // no comment lines in our file
            HasHeaders = true
        };

        // 5. Create a CsvDataSource from the CSV file.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 6. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // No special options required for this simple scenario.
        engine.BuildReport(loadedTemplate, csvDataSource, "persons");

        // 7. Save the generated report.
        loadedTemplate.Save(reportPath);

        // Inform the user (optional, no interactive input required).
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
