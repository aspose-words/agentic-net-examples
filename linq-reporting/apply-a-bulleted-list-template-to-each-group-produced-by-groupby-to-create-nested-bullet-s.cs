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

public class Group
{
    public string Category { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        List<Item> items = new()
        {
            new Item { Category = "Fruits", Name = "Apple" },
            new Item { Category = "Fruits", Name = "Banana" },
            new Item { Category = "Fruits", Name = "Orange" },
            new Item { Category = "Vegetables", Name = "Carrot" },
            new Item { Category = "Vegetables", Name = "Broccoli" },
            new Item { Category = "Grains", Name = "Rice" }
        };

        // Group items by Category.
        ReportModel model = new()
        {
            Groups = items
                .GroupBy(i => i.Category)
                .Select(g => new Group
                {
                    Category = g.Key,
                    Items = g.ToList()
                })
                .ToList()
        };

        // Create the LINQ Reporting template.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [g in model.Groups]>>");

        // Group bullet (level 1).
        builder.CurrentParagraph.ListFormat.ApplyBulletDefault();
        builder.CurrentParagraph.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[g.Category]>>");

        // Begin inner foreach over items.
        builder.Writeln("<<foreach [i in g.Items]>>");

        // Item bullet (level 2).
        builder.CurrentParagraph.ListFormat.ApplyBulletDefault();
        builder.CurrentParagraph.ListFormat.ListLevelNumber = 2;
        builder.Writeln("<<[i.Name]>>");

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
