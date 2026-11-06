using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code page support for CSV encoding.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Generate a sample CSV file.
        string csvPath = Path.Combine(outputDir, "sample.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Id,Name,Amount");
            for (int i = 1; i <= 1000; i++)
            {
                writer.WriteLine($"{i},Item {i},{i * 10.5}");
            }
        }

        // Build a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from CSV data:");
        builder.Writeln("<<foreach [row in rows]>>");
        builder.Writeln("<<[row.Id]>> - <<[row.Name]>> - <<[row.Amount]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure CSV loading options (default separator is comma).
        var csvLoadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Stream the CSV without loading it entirely into memory.
        using (var csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            var csvDataSource = new CsvDataSource(csvStream, csvLoadOptions);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, csvDataSource, "rows");
        }

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "ReportFromCsv.docx");
        doc.Save(outputPath);
    }
}
