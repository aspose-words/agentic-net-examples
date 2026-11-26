using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Record
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public double Value { get; set; }
}

public class ReportModel
{
    public List<Record> Records { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a folder for sample CSV files.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        // Create sample CSV files.
        for (int i = 1; i <= 3; i++)
        {
            string csvPath = Path.Combine(dataFolder, $"File{i}.csv");
            var lines = new List<string>
            {
                "Id,Name,Value",
                $"1,ItemA{i},10.5",
                $"2,ItemB{i},20.0",
                $"3,ItemC{i},30.25"
            };
            File.WriteAllLines(csvPath, lines, Encoding.UTF8);
        }

        // Load CSV files, parse them, and merge into a single collection.
        var model = new ReportModel();
        foreach (string csvFile in Directory.GetFiles(dataFolder, "*.csv"))
        {
            string[] allLines = File.ReadAllLines(csvFile, Encoding.UTF8);
            if (allLines.Length < 2) continue; // Skip if no data.

            // Assume first line is header.
            for (int lineIndex = 1; lineIndex < allLines.Length; lineIndex++)
            {
                string line = allLines[lineIndex];
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length != 3) continue;

                if (int.TryParse(parts[0], out int id) &&
                    double.TryParse(parts[2], out double value))
                {
                    var record = new Record
                    {
                        Id = id,
                        Name = parts[1],
                        Value = value
                    };
                    model.Records.Add(record);
                }
            }
        }

        // Create a LINQ Reporting template programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Merged CSV Report");
        builder.Writeln($"Generated on {DateTime.Now}");

        // Begin foreach loop over merged records.
        builder.Writeln("<<foreach [rec in Records]>>");

        // Create a simple table for each record.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[rec.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[rec.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[rec.Value]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
