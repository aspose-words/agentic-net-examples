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

        // Prepare sample CSV file with quoted fields containing commas.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        File.WriteAllText(csvPath,
            "Name,Description\r\n" +
            "\"John Doe\",\"Developer, C#\"\r\n" +
            "\"Jane Smith\",\"Manager, Sales\"\r\n");

        // Create a template Word document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Description: <<[row.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source options (default separator ',' and quote char '"').
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Create CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
