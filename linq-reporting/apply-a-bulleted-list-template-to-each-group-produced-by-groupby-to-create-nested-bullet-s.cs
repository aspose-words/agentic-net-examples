using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();

    // Groups items by Category for the report.
    public IEnumerable<IGrouping<string, Item>> Groups => Items.GroupBy(i => i.Category);
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items =
            {
                new Item { Category = "Fruits", Name = "Apple" },
                new Item { Category = "Fruits", Name = "Banana" },
                new Item { Category = "Fruits", Name = "Orange" },
                new Item { Category = "Vegetables", Name = "Carrot" },
                new Item { Category = "Vegetables", Name = "Broccoli" },
                new Item { Category = "Grains", Name = "Rice" },
                new Item { Category = "Grains", Name = "Wheat" }
            }
        };

        // Create the LINQ Reporting template.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report:");
        builder.Writeln();

        // Start a bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Outer foreach: groups.
        builder.Writeln("<<foreach [g in model.Groups]>><<[g.Key]>>");

        // Inner foreach: items within each group (second-level bullet).
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("<<foreach [item in g]>><<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Reset to first-level bullet for the next group.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<</foreach>>");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the final report.
        var outputPath = "Report.docx";
        templateDoc.Save(outputPath);
    }
}
