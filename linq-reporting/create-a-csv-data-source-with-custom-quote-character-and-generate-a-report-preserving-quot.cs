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

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample CSV file with custom quote character (single quote).
        string csvPath = Path.Combine(outputDir, "sample.csv");
        string[] csvLines =
        {
            "Id,Name,Description",
            "1,John,'Description with, a comma'",
            "2,Jane,'Another description, also with commas'"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Build a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Description: <<[row.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Configure CSV data load options with custom quote character.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            QuoteChar = '\'',
            HasHeaders = true
        };

        // Create CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, csvData, "data");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        template.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
