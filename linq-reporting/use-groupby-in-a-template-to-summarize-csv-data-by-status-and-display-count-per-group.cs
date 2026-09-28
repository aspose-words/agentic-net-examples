using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, @"Id,Name,Status
1,Apple,Available
2,Banana,OutOfStock
3,Cherry,Available
4,Date,Available
5,Elderberry,OutOfStock
6,Fig,Discontinued");

        // Load CSV into a list of records.
        List<Record> records = LoadCsv(csvPath);

        // Group records by Status.
        List<StatusGroup> groups = records
            .GroupBy(r => r.Status)
            .Select(g => new StatusGroup { Status = g.Key, Count = g.Count() })
            .ToList();

        // Prepare the model for the report.
        ReportModel model = new() { Groups = groups };

        // Create the template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Status Summary Report");
        builder.Writeln();
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Status: <<[g.Status]>>, Count: <<[g.Count]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        template.Save(outputPath);
    }

    private static List<Record> LoadCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        var records = new List<Record>();

        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length >= 3)
            {
                records.Add(new Record
                {
                    Id = int.Parse(parts[0]),
                    Name = parts[1],
                    Status = parts[2]
                });
            }
        }

        return records;
    }
}

public class Record
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}

public class StatusGroup
{
    public string Status { get; set; } = string.Empty;
    public int Count { get; set; }
}

public class ReportModel
{
    public List<StatusGroup> Groups { get; set; } = new();
}
