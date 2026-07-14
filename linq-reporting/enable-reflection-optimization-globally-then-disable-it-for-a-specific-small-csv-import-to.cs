using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Enable reflection optimization globally (default is true, but set explicitly).
        ReportingEngine.UseReflectionOptimization = true;

        // Create a small CSV file with headers.
        string csvPath = "people.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,25"
        });

        // Build a template document containing LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the saved template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Disable reflection optimization for this small CSV import to avoid overhead.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // Save the generated report.
        reportDoc.Save("report.docx");
    }
}
