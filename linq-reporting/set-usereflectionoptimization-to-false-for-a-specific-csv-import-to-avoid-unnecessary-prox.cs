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

        // Create sample CSV file.
        const string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25");

        // Create a template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Name: <<[row.Name]>> Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Load CSV data source with headers.
        var csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        var csvData = new CsvDataSource(csvPath, csvOptions);

        // Disable reflection optimization for this report.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "Data");

        // Save the generated report.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
