using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample CSV data
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Status\n" +
            "1,John Doe,Active\n" +
            "2,Jane Smith,Inactive\n" +
            "3,Bob Johnson,Active\n");

        // Load CSV rows into objects
        List<CsvRecord> allRecords = File.ReadAllLines(csvPath)
            .Skip(1) // skip header
            .Select(line => line.Split(','))
            .Where(parts => parts.Length >= 3)
            .Select(parts => new CsvRecord
            {
                Id = parts[0].Trim(),
                Name = parts[1].Trim(),
                Status = parts[2].Trim()
            })
            .ToList();

        // Filter rows where Status == "Active"
        List<CsvRecord> activeRecords = allRecords
            .Where(r => r.Status.Equals("Active", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Prepare the model for the report
        ReportModel model = new()
        {
            Records = activeRecords
        };

        // Create the template document programmatically
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Active Records Report");
        builder.Writeln("<<foreach [rec in Records]>>");

        // Table header
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Status");
        builder.EndRow();

        // Table rows (will be repeated by the foreach tag)
        builder.InsertCell();
        builder.Writeln("<<[rec.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[rec.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[rec.Status]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

public class CsvRecord
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

public class ReportModel
{
    public List<CsvRecord> Records { get; set; } = new();
}
