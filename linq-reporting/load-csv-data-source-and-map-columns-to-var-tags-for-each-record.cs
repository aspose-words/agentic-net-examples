using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // ReportingEngine, CsvDataSource, CsvDataLoadOptions
using Aspose.Words.Reporting; // Ensure correct namespace for CSV data source

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for non‑UTF8 encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a sample CSV file with headers and three rows.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Age\r\n" +
            "1,John Doe,30\r\n" +
            "2,Jane Smith,25\r\n" +
            "3,Bob Johnson,40");

        // Build a template document that contains LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the CSV rows (exposed as 'persons').
        builder.Writeln("<<foreach [person in persons]>>");
        // Output each column using tags that map to CSV headers.
        builder.Writeln("Id: <<[person.Id]>>, Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        var loadOptions = new CsvDataLoadOptions(true);
        // Create a CSV data source based on the file.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the data source. The root name must match the tag reference.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // Save the generated report.
        reportDoc.Save("report.docx");
    }
}
