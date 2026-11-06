using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Record
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

public class Model
{
    public List<Record> Records { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV reading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Status\n" +
            "1,Alpha,Active\n" +
            "2,Beta,Inactive\n" +
            "3,Gamma,Active\n" +
            "4,Delta,Inactive\n");

        // Load CSV rows into a list of Record objects.
        List<Record> allRecords = new();
        foreach (var line in File.ReadAllLines(csvPath).Skip(1))
        {
            var parts = line.Split(',');
            if (parts.Length != 3) continue;
            if (!int.TryParse(parts[0], out int id)) continue;
            allRecords.Add(new Record
            {
                Id = id,
                Name = parts[1],
                Status = parts[2]
            });
        }

        // Filter rows where Status equals "Active".
        List<Record> activeRecords = allRecords
            .Where(r => r.Status.Equals("Active", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create the model with filtered records.
        Model model = new() { Records = activeRecords };

        // Build the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert a simple foreach loop to list records.
        builder.Writeln("<<foreach [rec in Records]>>");
        builder.Writeln("Id: <<[rec.Id]>>, Name: <<[rec.Name]>>, Status: <<[rec.Status]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report using the filtered model.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
