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
            new Item { Category = "Grains", Name = "Rice" }
        };

        // Group items by Category.
        List<Group> groups = items
            .GroupBy(i => i.Category)
            .Select(g => new Group { Category = g.Key, Items = g.ToList() })
            .ToList();

        // Build the wrapper model for the report.
        ReportModel model = new() { Groups = groups };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list that will be used for both group headers and items.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [g in Groups]>>");

        // Group header – level 0 bullet.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[g.Category]>>");

        // Begin inner foreach over items within the current group.
        builder.Writeln("<<foreach [i in g.Items]>>");

        // Item name – level 1 bullet.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[i.Name]>>");

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer foreach.
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
