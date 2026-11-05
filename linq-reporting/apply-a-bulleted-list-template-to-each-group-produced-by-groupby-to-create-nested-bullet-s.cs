using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class GroupModel
{
    public string Key { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<GroupModel> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var items = new List<Item>
        {
            new() { Name = "Apple", Category = "Fruits" },
            new() { Name = "Banana", Category = "Fruits" },
            new() { Name = "Carrot", Category = "Vegetables" },
            new() { Name = "Broccoli", Category = "Vegetables" },
            new() { Name = "Chicken", Category = "Meat" }
        };

        // Group items by Category.
        var groups = items
            .GroupBy(i => i.Category)
            .Select(g => new GroupModel { Key = g.Key, Items = g.ToList() })
            .ToList();

        var model = new ReportModel { Groups = groups };

        // Create the LINQ Reporting template.
        string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report:");
        builder.Writeln();

        // Begin group foreach.
        builder.Writeln("<<foreach [group in Groups]>>");

        // Group bullet.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[group.Key]>>");

        // Nested items bullet.
        builder.ListFormat.ListIndent();
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.ListFormat.ListOutdent();

        // End group foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string outputPath = "ReportOutput.docx";
        templateDoc.Save(outputPath);
    }
}
