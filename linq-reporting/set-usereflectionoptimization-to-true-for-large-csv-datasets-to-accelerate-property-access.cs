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
        string[] csvLines =
        {
            "Id,Name,Age",
            "1,John Doe,30",
            "2,Jane Smith,25",
            "3,Bob Johnson,40"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template document for report generation.
        Document reportDoc = new Document(templatePath);

        // Load CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, new CsvDataLoadOptions { HasHeaders = true });

        // Enable reflection optimization for large data sets.
        ReportingEngine.UseReflectionOptimization = true;

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, csvDataSource, "data");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
