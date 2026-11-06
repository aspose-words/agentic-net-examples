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
        const string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Age\r\n" +
            "1,John Doe,30\r\n" +
            "2,Jane Smith,25\r\n" +
            "3,Bob Johnson,40");

        // Create a Word template with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over the CSV records (named "Records").
        builder.Writeln("<<foreach [rec in Records]>>");
        // Map CSV columns to <<var>> tags.
        builder.Writeln("Id: <<[rec.Id]>>");
        builder.Writeln("Name: <<[rec.Name]>>");
        builder.Writeln("Age: <<[rec.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Load CSV data source with headers.
        var csvData = new CsvDataSource(csvPath, new CsvDataLoadOptions { HasHeaders = true });

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "Records");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
