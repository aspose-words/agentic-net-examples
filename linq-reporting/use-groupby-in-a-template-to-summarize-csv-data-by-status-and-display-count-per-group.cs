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
        // Register code page provider for CSV reading (if needed)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, @"Id,Name,Status
1,Apple,Available
2,Banana,OutOfStock
3,Cherry,Available
4,Date,Discontinued
5,Elderberry,Available
6,Fig,OutOfStock");

        // Load CSV into objects
        List<Item> items = LoadCsv(csvPath);

        // Group by Status
        List<GroupInfo> groups = items
            .GroupBy(i => i.Status)
            .Select(g => new GroupInfo { Status = g.Key, Count = g.Count() })
            .ToList();

        // Prepare the model
        ReportModel model = new()
        {
            Groups = groups
        };

        // Create the template document
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Summary by Status:");
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Status: <<[g.Status]>>, Count: <<[g.Count]>>");
        builder.Writeln("<</foreach>>");

        // Build the report
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // Save the result
        string outputPath = "output.docx";
        template.Save(outputPath);
    }

    private static List<Item> LoadCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        var items = new List<Item>();

        // Assume first line contains headers
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length >= 3)
            {
                items.Add(new Item
                {
                    Id = int.Parse(parts[0]),
                    Name = parts[1],
                    Status = parts[2]
                });
            }
        }

        return items;
    }
}

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}

public class GroupInfo
{
    public string Status { get; set; } = string.Empty;
    public int Count { get; set; }
}

public class ReportModel
{
    public List<GroupInfo> Groups { get; set; } = new();
}
