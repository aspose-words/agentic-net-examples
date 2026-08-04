using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class StatusGroup
{
    public string Status { get; set; } = string.Empty;
    public int Count { get; set; }
}

public class ReportModel
{
    public List<StatusGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        CreateSampleCsv(csvPath);

        // Load CSV data, group by Status, and build the report model.
        ReportModel model = BuildReportModelFromCsv(csvPath);

        // Create a LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Build the report using the model.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "ReportByStatus.docx";
        doc.Save(outputPath);
    }

    // Generates a simple CSV file with Id, Name, and Status columns.
    private static void CreateSampleCsv(string path)
    {
        var lines = new[]
        {
            "Id,Name,Status",
            "1,Alpha,Open",
            "2,Beta,Closed",
            "3,Gamma,Open",
            "4,Delta,InProgress",
            "5,Epsilon,Closed",
            "6,Zeta,Open"
        };
        File.WriteAllLines(path, lines, Encoding.UTF8);
    }

    // Reads the CSV file, groups records by the Status column, and returns a populated model.
    private static ReportModel BuildReportModelFromCsv(string csvPath)
    {
        var lines = File.ReadAllLines(csvPath, Encoding.UTF8);
        if (lines.Length < 2)
            return new ReportModel();

        // Parse header to find column indexes.
        var headers = lines[0].Split(',');
        int statusIndex = Array.IndexOf(headers, "Status");
        if (statusIndex < 0)
            throw new InvalidOperationException("CSV does not contain a 'Status' column.");

        // Extract statuses from data rows.
        var statuses = lines
            .Skip(1)
            .Select(line => line.Split(',')[statusIndex])
            .Where(status => !string.IsNullOrWhiteSpace(status));

        // Group by status and create model groups.
        var groups = statuses
            .GroupBy(s => s)
            .Select(g => new StatusGroup { Status = g.Key, Count = g.Count() })
            .OrderBy(g => g.Status)
            .ToList();

        return new ReportModel { Groups = groups };
    }

    // Creates a Word document containing LINQ Reporting tags that iterate over the Groups collection.
    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Summary of records grouped by Status:");
        builder.Writeln("<<foreach [group in model.Groups]>>");
        builder.Writeln("Status: <<[group.Status]>>, Count: <<[group.Count]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}
