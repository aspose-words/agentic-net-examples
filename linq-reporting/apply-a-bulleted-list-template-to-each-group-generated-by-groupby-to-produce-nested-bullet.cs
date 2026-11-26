using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
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
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
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
            .Select(g => new Group { Category = g.Key, Items = g.ToList() })
            .ToList();

        var model = new ReportModel { Groups = groups };

        // Create the LINQ Reporting template.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report of Items grouped by Category:");
        builder.Writeln();

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [g in Groups]>>");
        // Bullet for group name.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[g.Category]>>");
        // Increase list level for items.
        builder.ListFormat.ListIndent();

        // Begin inner foreach over items in the current group.
        builder.Writeln("<<foreach [i in g.Items]>>");
        builder.Writeln("<<[i.Name]>>");
        builder.Writeln("<</foreach>>");

        // Restore list level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
