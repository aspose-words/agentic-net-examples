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
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "sampleData.json";
        var sampleItems = new List<Item>
        {
            new Item { Name = "Apple", Category = "Fruit" },
            new Item { Name = "Carrot", Category = "Vegetable" },
            new Item { Name = "Banana", Category = "fruit" },
            new Item { Name = "Broccoli", Category = "Vegetable" }
        };
        var wrapper = new Wrapper { Items = sampleItems };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(wrapper, Formatting.Indented));

        // Load JSON into objects.
        string jsonContent = File.ReadAllText(jsonPath);
        var data = JsonConvert.DeserializeObject<Wrapper>(jsonContent) ?? new Wrapper();

        // Filter items where Category equals "fruit" (case‑insensitive).
        var filtered = data.Items
            .Where(i => string.Equals(i.Category, "fruit", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Use the filtered collection as the data source for the report.
        var reportModel = new Wrapper { Items = filtered };

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Filtered Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(" - <<[item.Name]>> (<<[item.Category]>>)");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, reportModel, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        templateDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Wrapper class exposing a collection of items.
public class Wrapper
{
    public List<Item> Items { get; set; } = new();
}

// Data model for each JSON entry.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}
