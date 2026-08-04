using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with boolean values.
        string csvPath = "people.csv";
        string[] csvLines =
        {
            "Name,IsActive",
            "Alice,true",
            "Bob,false",
            "Charlie,true"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Create a simple Word template containing LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use a foreach block to iterate over the CSV rows.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Active: <<[p.IsActive]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template document.
        Document reportDoc = new Document(templatePath);

        // Configure CSV load options to treat the first row as headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };

        // Create a CSV data source using the configured options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "persons");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine("Report generated: " + Path.GetFullPath(outputPath));
    }
}
