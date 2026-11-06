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

        // Create a CSV file with quoted fields that contain commas.
        string csvPath = "sample.csv";
        string[] csvLines =
        {
            "\"Name\",\"Address\",\"Notes\"",
            "\"John Doe\",\"123 Main St, Apt 4\",\"Loves, commas\"",
            "\"Jane Smith\",\"456 Oak Ave\",\"No commas\""
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Build a Word template that uses LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("CSV Report");
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Address: <<[row.Address]>>");
        builder.Writeln("Notes: <<[row.Notes]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template document.
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options to respect quoted fields.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            QuoteChar = '"',
            HasHeaders = true
        };

        // Load the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, csvOptions);

        // Generate the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "Data");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
