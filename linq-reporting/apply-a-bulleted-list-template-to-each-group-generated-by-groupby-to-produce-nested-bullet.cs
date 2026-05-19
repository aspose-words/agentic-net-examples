using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        List<Item> items = new()
        {
            new Item { Category = "Fruits", Name = "Apple" },
            new Item { Category = "Fruits", Name = "Banana" },
            new Item { Category = "Fruits", Name = "Cherry" },
            new Item { Category = "Vegetables", Name = "Carrot" },
            new Item { Category = "Vegetables", Name = "Lettuce" },
            new Item { Category = "Grains", Name = "Rice" },
            new Item { Category = "Grains", Name = "Wheat" }
        };

        // Group items by Category.
        List<Group> groups = items
            .GroupBy(i => i.Category)
            .Select(g => new Group { Key = g.Key, Items = g.ToList() })
            .ToList();

        // Wrap groups into a root model.
        ReportModel model = new() { Groups = groups };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list style.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // top‑level bullets for groups.

        // Outer foreach – iterate over groups.
        builder.Writeln("<<foreach [group in Groups]>>");
        // Group header (top‑level bullet).
        builder.Writeln("<<[group.Key]>>");

        // Switch to second list level for items inside each group.
        builder.ListFormat.ListLevelNumber = 1;

        // Inner foreach – iterate over items of the current group.
        builder.Writeln("<<foreach [item in group.Items]>>");
        // Item name (second‑level bullet).
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Return to top‑level for the next group header.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}

public class Group
{
    public string Key { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}
