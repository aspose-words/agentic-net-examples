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
        File.WriteAllText(csvPath, "Name,Age\r\nAlice,30\r\nBob,25\r\nCharlie,35", Encoding.UTF8);

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer List:");
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Load CSV data source.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Disable reflection optimization for this small CSV processing.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "Data");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
