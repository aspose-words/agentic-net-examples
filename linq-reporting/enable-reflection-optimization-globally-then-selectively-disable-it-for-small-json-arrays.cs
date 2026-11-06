using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare sample data.
        var largeData = new DataWrapper
        {
            Items = new List<Item>()
        };
        for (int i = 1; i <= 100; i++)
        {
            largeData.Items.Add(new Item { Id = i, Name = $"Item {i}" });
        }

        var smallData = new DataWrapper
        {
            Items = new List<Item>
            {
                new Item { Id = 1, Name = "First" },
                new Item { Id = 2, Name = "Second" }
            }
        };

        // Serialize data to JSON files (kept for demonstration purposes).
        File.WriteAllText("large.json", JsonConvert.SerializeObject(largeData));
        File.WriteAllText("small.json", JsonConvert.SerializeObject(smallData));

        // Create a template for the large data report.
        var largeTemplate = new Document();
        var largeBuilder = new DocumentBuilder(largeTemplate);
        largeBuilder.Writeln("Large Data Report");
        largeBuilder.Writeln("<<foreach [item in json.Items]>>");
        largeBuilder.Writeln("<<[item.Id]>> - <<[item.Name]>>");
        largeBuilder.Writeln("<</foreach>>");
        largeTemplate.Save("LargeTemplate.docx");

        // Create a template for the small data report.
        var smallTemplate = new Document();
        var smallBuilder = new DocumentBuilder(smallTemplate);
        smallBuilder.Writeln("Small Data Report");
        smallBuilder.Writeln("<<foreach [item in json.Items]>>");
        smallBuilder.Writeln("<<[item.Id]>> - <<[item.Name]>>");
        smallBuilder.Writeln("<</foreach>>");
        smallTemplate.Save("SmallTemplate.docx");

        // Build report for large data (reflection optimization remains enabled).
        var largeDoc = new Document("LargeTemplate.docx");
        var largeEngine = new ReportingEngine();
        largeEngine.BuildReport(largeDoc, largeData, "json");
        largeDoc.Save("LargeReport.docx", SaveFormat.Docx);

        // Disable reflection optimization for small data to balance performance.
        ReportingEngine.UseReflectionOptimization = false;

        // Build report for small data.
        var smallDoc = new Document("SmallTemplate.docx");
        var smallEngine = new ReportingEngine();
        smallEngine.BuildReport(smallDoc, smallData, "json");
        smallDoc.Save("SmallReport.docx", SaveFormat.Docx);

        // Re‑enable reflection optimization for subsequent operations if needed.
        ReportingEngine.UseReflectionOptimization = true;
    }
}

// Data model used for JSON serialization and reporting.
public class DataWrapper
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}
