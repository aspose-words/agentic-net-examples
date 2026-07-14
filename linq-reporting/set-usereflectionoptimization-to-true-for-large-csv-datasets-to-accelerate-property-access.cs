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

        // ---------- Create a sample CSV file ----------
        const string csvFile = "persons.csv";
        using (var writer = new StreamWriter(csvFile))
        {
            // Header row
            writer.WriteLine("Name,Age,Country");
            // Generate 100 sample records
            for (int i = 1; i <= 100; i++)
            {
                writer.WriteLine($"Person{i},{20 + i % 30},Country{i % 5}");
            }
        }

        // ---------- Build a template document with LINQ Reporting tags ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Person List:");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("Country: <<[p.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, demonstrates the load step later).
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template document back.
        var loadedTemplate = new Document(templatePath);

        // ---------- Prepare CSV data source ----------
        var csvLoadOptions = new CsvDataLoadOptions(true); // CSV has headers.
        var csvDataSource = new CsvDataSource(csvFile, csvLoadOptions);

        // ---------- Enable reflection optimization for large data sets ----------
        ReportingEngine.UseReflectionOptimization = true;

        // ---------- Build the report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, csvDataSource, "persons");

        // ---------- Save the generated report ----------
        const string outputPath = "report.docx";
        loadedTemplate.Save(outputPath);
    }
}
