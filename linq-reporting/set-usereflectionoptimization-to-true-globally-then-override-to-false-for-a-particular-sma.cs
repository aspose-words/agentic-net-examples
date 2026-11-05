using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Simple header showing the report title.
        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln();

        // Iterate over a collection named Items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(" - Item: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // ---------- Large data source (uses global reflection optimization) ----------
        var largeData = new ReportModel
        {
            Title = "Large Data Report",
            Items = new List<Item>()
        };

        for (int i = 1; i <= 100; i++)
        {
            largeData.Items.Add(new Item { Name = $"LargeItem{i}" });
        }

        // Load the template and build the report for the large data source.
        Document largeDoc = new Document(templatePath);
        ReportingEngine largeEngine = new ReportingEngine();
        largeEngine.BuildReport(largeDoc, largeData, "model");
        largeDoc.Save("LargeReport.docx");

        // ---------- Small data source (override reflection optimization) ----------
        // Temporarily disable reflection optimization for this small data source.
        ReportingEngine.UseReflectionOptimization = false;

        var smallData = new ReportModel
        {
            Title = "Small Data Report",
            Items = new List<Item>
            {
                new Item { Name = "SmallItemA" },
                new Item { Name = "SmallItemB" }
            }
        };

        // Load the template and build the report for the small data source.
        Document smallDoc = new Document(templatePath);
        ReportingEngine smallEngine = new ReportingEngine();
        smallEngine.BuildReport(smallDoc, smallData, "model");
        smallDoc.Save("SmallReport.docx");

        // Restore the global setting if needed.
        ReportingEngine.UseReflectionOptimization = true;
    }
}

// Root data model used by both reports.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Simple item class.
public class Item
{
    public string Name { get; set; } = string.Empty;
}
