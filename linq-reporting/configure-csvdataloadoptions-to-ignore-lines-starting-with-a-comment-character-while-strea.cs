using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support (required for some code pages).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a sample CSV file that contains comment lines starting with '#'.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("# This is a comment and should be ignored");
            writer.WriteLine("# Another comment line");
            writer.WriteLine("Name,Age,Country"); // Header
            writer.WriteLine("Alice,30,USA");
            writer.WriteLine("Bob,25,Canada");
            writer.WriteLine("# Comment between data rows");
            writer.WriteLine("Charlie,35,UK");
        }

        // Configure CSV loading options.
        // Only the properties that exist in the current Aspose.Words version are set.
        // CommentChar tells the loader to skip lines that start with '#'.
        var csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            CommentChar = '#'
            // Separator and Encoding use default values (',' and UTF‑8) which are suitable for this example.
        };

        // Create a CSV data source using the configured options.
        var csvDataSource = new CsvDataSource(csvPath, csvOptions);

        // Build a simple Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("CSV Report");
        builder.Writeln();

        // Foreach loop over CSV rows (root name will be "data").
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Age: <<[row.Age]>>");
        builder.Writeln("Country: <<[row.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using the CSV data source.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
