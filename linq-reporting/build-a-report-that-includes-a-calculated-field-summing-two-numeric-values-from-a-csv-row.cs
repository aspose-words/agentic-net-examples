using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingCsvSum
{
    // Model representing a CSV row.
    public class CsvRow
    {
        public string Item { get; set; } = string.Empty;
        public int Value1 { get; set; }
        public int Value2 { get; set; }

        // Calculated field – sum of the two numeric values.
        public int Sum => Value1 + Value2;
    }

    public static void Main()
    {
        // Register code page provider for CSV handling.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Define file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string csvPath = Path.Combine(outputDir, "data.csv");
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // Create sample CSV data with headers and numeric values.
        File.WriteAllLines(csvPath, new[]
        {
            "Item,Value1,Value2",
            "A,10,15",
            "B,20,5",
            "C,7,13"
        });

        // Build the template document containing LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from CSV data:");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Item: <<[row.Item]>>");
        builder.Writeln("Value1: <<[row.Value1]>>");
        builder.Writeln("Value2: <<[row.Value2]>>");
        // Calculated field: sum of Value1 and Value2.
        builder.Writeln("Sum: <<[row.Sum]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Parse the CSV file into a list of strongly‑typed objects.
        List<CsvRow> rows = new List<CsvRow>();
        string[] csvLines = File.ReadAllLines(csvPath);
        // Skip header line (index 0).
        for (int i = 1; i < csvLines.Length; i++)
        {
            string[] parts = csvLines[i].Split(',');
            if (parts.Length != 3) continue; // Guard against malformed lines.

            rows.Add(new CsvRow
            {
                Item = parts[0],
                Value1 = int.TryParse(parts[1], out int v1) ? v1 : 0,
                Value2 = int.TryParse(parts[2], out int v2) ? v2 : 0
            });
        }

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name must match the name used in the template tags ("data").
        engine.BuildReport(reportDoc, rows, "data");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
