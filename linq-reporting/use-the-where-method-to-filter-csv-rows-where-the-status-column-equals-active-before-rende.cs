using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Record
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

public class ReportModel
{
    public List<Record> Records { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create sample CSV data.
        const string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Status\n" +
            "1,Alpha,Active\n" +
            "2,Beta,Inactive\n" +
            "3,Gamma,Active\n" +
            "4,Delta,Inactive\n");

        // Load CSV into a list of Record objects.
        List<Record> allRecords = new();
        foreach (var line in File.ReadAllLines(csvPath).Skip(1))
        {
            var parts = line.Split(',');
            if (parts.Length != 3) continue;
            allRecords.Add(new Record
            {
                Id = int.Parse(parts[0]),
                Name = parts[1],
                Status = parts[2]
            });
        }

        // Filter rows where Status equals "Active".
        List<Record> activeRecords = allRecords
            .Where(r => r.Status.Equals("Active", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Prepare the model for the report.
        ReportModel model = new() { Records = activeRecords };

        // Create the LINQ Reporting template.
        const string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Active Records:");
        builder.Writeln("<<foreach [rec in Records]>>");
        builder.Writeln("Id: <<[rec.Id]>>, Name: <<[rec.Name]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
