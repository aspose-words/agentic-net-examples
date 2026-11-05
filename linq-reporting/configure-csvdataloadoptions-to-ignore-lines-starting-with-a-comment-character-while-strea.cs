using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Row
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Country { get; set; } = "";
}

public class ReportModel
{
    // Expose the CSV rows to the template.
    public List<Row> Data { get; } = new();

    public ReportModel(List<Row> data) => Data = data;
}

public class Program
{
    public static void Main()
    {
        // Enable support for code pages (required for some CSV encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create a sample CSV file that contains comment lines starting with '#'.
        // -----------------------------------------------------------------
        string csvPath = "sample.csv";
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("# This is a comment line and should be ignored");
            writer.WriteLine("# Another comment");
            writer.WriteLine("Name,Age,Country");
            writer.WriteLine("Alice,30,USA");
            writer.WriteLine("Bob,25,Canada");
            writer.WriteLine("# Comment after data");
            writer.WriteLine("Charlie,35,UK");
        }

        // -----------------------------------------------------------------
        // Parse the CSV file while ignoring comment lines that start with '#'.
        // This mimics the behavior of CsvDataLoadOptions.CommentCharEnabled.
        // -----------------------------------------------------------------
        var rows = new List<Row>();
        using (var reader = new StreamReader(csvPath, Encoding.UTF8))
        {
            string? line;
            string[]? headers = null;

            while ((line = reader.ReadLine()) != null)
            {
                // Skip empty lines and comment lines.
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                    continue;

                // The first non‑comment line is assumed to be the header.
                if (headers == null)
                {
                    headers = line.Split(',');
                    continue;
                }

                // Parse data rows.
                var fields = line.Split(',');
                if (fields.Length != headers.Length)
                    continue; // Skip malformed rows.

                var row = new Row
                {
                    Name = fields[0],
                    Age = int.TryParse(fields[1], out var age) ? age : 0,
                    Country = fields[2]
                };
                rows.Add(row);
            }
        }

        // -----------------------------------------------------------------
        // Build a Word template that uses LINQ Reporting tags to iterate over rows.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from CSV data:");
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>, Country: <<[row.Country]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Wrap the parsed rows in a model class that matches the template root name.
        var model = new ReportModel(rows);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
