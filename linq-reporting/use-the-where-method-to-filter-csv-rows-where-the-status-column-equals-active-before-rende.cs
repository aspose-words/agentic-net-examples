using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // Create sample CSV file
        string csvPath = Path.Combine(workDir, "sample.csv");
        File.WriteAllText(csvPath,
            "Id,Name,Status\n" +
            "1,John Doe,Active\n" +
            "2,Jane Smith,Inactive\n" +
            "3,Bob Johnson,Active\n");

        // Load CSV rows into objects
        List<CsvRow> allRows = new();
        foreach (var line in File.ReadAllLines(csvPath).Skip(1)) // skip header
        {
            var parts = line.Split(',');
            if (parts.Length >= 3)
            {
                allRows.Add(new CsvRow
                {
                    Id = parts[0],
                    Name = parts[1],
                    Status = parts[2]
                });
            }
        }

        // Filter rows where Status == "Active" using LINQ Where
        List<CsvRow> activeRows = allRows
            .Where(r => string.Equals(r.Status, "Active", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Prepare the model for reporting
        ReportModel model = new()
        {
            Rows = activeRows
        };

        // Create a Word template programmatically
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Active Users Report");
        builder.Writeln("<<foreach [row in Rows]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Status: <<[row.Status]>>");
        builder.Writeln("<</foreach>>");

        // Build the report
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(workDir, "Report.docx");
        doc.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<CsvRow> Rows { get; set; } = new();
}

public class CsvRow
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}
