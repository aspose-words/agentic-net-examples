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

        // Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllText(csvPath,
            "Name,Age\r\n" +
            "John Doe,30\r\n" +
            ",\r\n" + // Empty row – will produce empty paragraphs.
            "Jane Smith,25\r\n");

        // Create a template document programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Static section – will remain unchanged.
        builder.Writeln("=== Report Header ===");
        builder.Writeln();

        // CSV‑driven section.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        var loadOptions = new CsvDataLoadOptions(true);
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Set up the reporting engine to remove empty paragraphs after processing.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the CSV data source.
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the final document.
        doc.Save("Report_Output.docx");
    }
}
