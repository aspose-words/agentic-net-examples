using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible CSV encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, @"Id,Name,Category,Value
1,Alpha,A,100
2,Beta,B,200
3,Gamma,A,150
4,Delta,C,300
5,Epsilon,A,120");

        // Load CSV rows into objects.
        List<CsvRow> allRows = new();
        foreach (var line in File.ReadAllLines(csvPath, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("Id,"))
                continue; // Skip header or empty lines.

            string[] parts = line.Split(',');
            if (parts.Length != 4)
                continue;

            allRows.Add(new CsvRow
            {
                Id = parts[0],
                Name = parts[1],
                Category = parts[2],
                Value = parts[3]
            });
        }

        // Filter rows where Category == "A".
        ReportModel model = new()
        {
            Items = allRows.FindAll(r => r.Category == "A")
        };

        // Create a Word template programmatically.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Filtered Items (Category = A):");
        builder.ListFormat.ApplyNumberDefault();

        // Restart numbering and start foreach loop.
        builder.Writeln("<<restartNum>><<foreach [item in Items]>>");
        // Custom formatting: blue text for the name.
        builder.Writeln("<<textColor [\"Blue\"]>><<[item.Name]>> <</textColor>> - <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        builder.ListFormat.RemoveNumbers();

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model for a CSV row.
public class CsvRow
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

// Wrapper model passed to the reporting engine.
public class ReportModel
{
    public List<CsvRow> Items { get; set; } = new();
}
