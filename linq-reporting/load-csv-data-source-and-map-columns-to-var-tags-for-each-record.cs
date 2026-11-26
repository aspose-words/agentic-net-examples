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
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "Id,Name,Age\n1,John Doe,30\n2,Jane Smith,25\n3,Bob Johnson,40");

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [rec in csv]>>");
        builder.Writeln("Id: <<[rec.Id]>>");
        builder.Writeln("Name: <<[rec.Name]>>");
        builder.Writeln("Age: <<[rec.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Load CSV data source with headers.
        var csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        var csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "csv");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
