using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new List<Item>
        {
            new Item { Title = "Review project plan", Category = "Important" },
            new Item { Title = "Buy office supplies", Category = "Routine" },
            new Item { Title = "Prepare presentation", Category = "Important" },
            new Item { Title = "Team lunch", Category = "Social" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON and filter.
        var allItems = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new List<Item>();
        var filteredItems = allItems.Where(i => i.Category == "Important").ToList();

        // 3. Prepare the model for reporting.
        var model = new ReportModel
        {
            Items = filteredItems
        };

        // 4. Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Bullet‑point Summary:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("• <<[item.Title]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 5. Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated report.
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model for individual items.
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}

// Wrapper model used as the root object for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}
