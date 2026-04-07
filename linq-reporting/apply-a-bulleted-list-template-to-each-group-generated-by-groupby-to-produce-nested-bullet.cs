using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
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
        var items = new List<Item>
        {
            new() { Category = "Fruits", Name = "Apple" },
            new() { Category = "Fruits", Name = "Banana" },
            new() { Category = "Vegetables", Name = "Carrot" },
            new() { Category = "Vegetables", Name = "Lettuce" }
        };

        // Group items by Category.
        var groups = items
            .GroupBy(i => i.Category)
            .Select(g => new Group { Category = g.Key, Items = g.ToList() })
            .ToList();

        var model = new ReportModel { Groups = groups };

        // Paths for template and output.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a bullet list style for outer and inner levels.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // Outer bullet (group name).
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // first level
        builder.Writeln("<<[group.Category]>>");

        // Inner bullet (items within the group).
        builder.ListFormat.ListLevelNumber = 1; // second level
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Remove list formatting after the loop.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // The root object name is "model" because the template references "Groups".
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(resultPath);
    }
}
