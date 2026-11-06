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

        // Create a small CSV file with sample data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25");

        // Build a template document containing LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a foreach loop that iterates over the data source named "Data".
        builder.Writeln("<<foreach [p in Data]>>");
        builder.Writeln("Name: <<[p[\"Name\"]]>>, Age: <<[p[\"Age\"]]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Disable reflection optimization for a single small CSV file.
        ReportingEngine.UseReflectionOptimization = false;

        // Configure CSV loading options to treat the first row as headers.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Create a CSV data source using the configured options.
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "Data");

        // Save the generated report.
        reportDoc.Save("report.docx");
    }
}
