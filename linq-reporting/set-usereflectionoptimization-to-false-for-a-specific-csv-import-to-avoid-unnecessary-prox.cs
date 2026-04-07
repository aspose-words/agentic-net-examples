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
        string csvContent = @"Id,Name,Age
1,John Doe,30
2,Jane Smith,25
3,Bob Johnson,40";
        File.WriteAllText(csvPath, csvContent, Encoding.UTF8);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [row in CsvData]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Load CSV data source.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Disable reflection optimization for this import.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "CsvData");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
