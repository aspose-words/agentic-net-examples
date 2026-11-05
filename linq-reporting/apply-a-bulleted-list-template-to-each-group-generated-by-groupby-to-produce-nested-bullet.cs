using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var rawItems = new List<RawItem>
        {
            new() { Category = "Fruits", Name = "Apple" },
            new() { Category = "Fruits", Name = "Banana" },
            new() { Category = "Fruits", Name = "Orange" },
            new() { Category = "Vegetables", Name = "Carrot" },
            new() { Category = "Vegetables", Name = "Broccoli" },
            new() { Category = "Grains", Name = "Rice" },
            new() { Category = "Grains", Name = "Wheat" }
        };

        var model = new ReportModel
        {
            Groups = rawItems
                .GroupBy(r => r.Category)
                .Select(g => new Group
                {
                    Category = g.Key,
                    Items = g.Select(i => new Item { Name = i.Name }).ToList()
                })
                .ToList()
        };

        // Create the LINQ Reporting template.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        // Start outer foreach over groups.
        builder.Writeln("<<foreach [g in Groups]>>");
        // Apply bullet for group name.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[g.Category]>>");
        // Indent for nested items.
        builder.ListFormat.ListIndent();
        // Start inner foreach over items.
        builder.Writeln("<<foreach [i in g.Items]>>");
        // Bullet for each item (inherits list format).
        builder.Writeln("<<[i.Name]>>");
        builder.Writeln("<</foreach>>");
        // Outdent back to outer level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Category { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
}

// Helper class for initial raw data.
public class RawItem
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
}
