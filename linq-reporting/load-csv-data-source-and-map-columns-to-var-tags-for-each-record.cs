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
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Age\r\n" +
            "1,John Doe,30\r\n" +
            "2,Jane Smith,25\r\n" +
            "3,Bob Johnson,40\r\n",
            Encoding.UTF8);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [rec in Records]>>");
        builder.Writeln("Id: <<[rec.Id]>>");
        builder.Writeln("Name: <<[rec.Name]>>");
        builder.Writeln("Age: <<[rec.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template.
        Document reportDoc = new Document(templatePath);

        // Load CSV data source (first row contains headers).
        var csvLoadOptions = new CsvDataLoadOptions { HasHeaders = true };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvLoadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };
        bool success = engine.BuildReport(reportDoc, csvData, "Records");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
