using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the working directory.
        const string templatePath = "Template.docx";
        const string csvPath = "Data.csv";
        const string outputPath = "Report.docx";

        // Create a CSV file with quoted fields that contain commas.
        // The first line contains headers.
        string csvContent =
            "\"Name\",\"Description\"\n" +
            "\"Item 1\",\"This, is a description with, commas\"\n" +
            "\"Item 2\",\"Another description, with commas\"\n";
        File.WriteAllText(csvPath, csvContent, Encoding.UTF8);

        // Build a simple Word template that uses LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from CSV data:");
        builder.Writeln("<<foreach [row in persons]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Description: <<[row.Description]>>");
        builder.Writeln("<</foreach>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back (simulating a separate load step).
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options to handle quoted fields and headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            HasHeaders = true,
            QuoteChar = '"', // Ensure quotes are recognized.
            Delimiter = ',', // Default delimiter, set explicitly for clarity.
            CommentChar = '\0' // No comment character.
        };

        // Create the CSV data source.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "persons");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
