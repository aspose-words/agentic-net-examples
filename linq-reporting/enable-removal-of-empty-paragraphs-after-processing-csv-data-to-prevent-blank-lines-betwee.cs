using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25\nCharlie,35");

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over CSV rows.
        builder.Writeln("<<foreach [row in data]>>");
        // Output each row in its own paragraph.
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Create CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, new CsvDataLoadOptions
        {
            // Use first row as header.
            HasHeaders = true
            // Default separator is a comma, so no explicit Separator property is needed.
        });

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the CSV data source.
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
