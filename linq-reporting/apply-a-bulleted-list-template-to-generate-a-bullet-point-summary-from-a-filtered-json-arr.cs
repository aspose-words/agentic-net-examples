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
        // Register code page provider for possible legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Step 1: Prepare sample JSON data ----------
        string jsonPath = "sample.json";
        var sampleItems = new List<Item>
        {
            new() { Title = "Breaking News: Market Rally", Category = "News" },
            new() { Title = "Weekly Sports Recap", Category = "Sports" },
            new() { Title = "Tech Insights: AI Trends", Category = "News" },
            new() { Title = "Cooking Tips: Summer Salads", Category = "Lifestyle" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleItems, Formatting.Indented));

        // Load JSON and filter for the "News" category.
        var allItems = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new();
        var filteredItems = allItems.Where(i => i.Category == "News").ToList();

        // ---------- Step 2: Create the LINQ Reporting template ----------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title paragraph.
        builder.Writeln("News Summary");
        builder.Writeln();

        // Apply bullet list formatting for the upcoming list.
        builder.ListFormat.ApplyBulletDefault();

        // Insert foreach tag that iterates over the filtered items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</foreach>>");

        // Remove bullet formatting after the list.
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------- Step 3: Load the template and build the report ----------
        var reportDoc = new Document(templatePath);
        var model = new ReportModel { Items = filteredItems };

        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // ---------- Step 4: Save the generated report ----------
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}

// Public data model representing each item.
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}

// Wrapper model passed as the root object to the reporting engine.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}
