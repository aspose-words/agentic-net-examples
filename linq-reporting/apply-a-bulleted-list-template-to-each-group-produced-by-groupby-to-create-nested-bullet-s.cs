using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Category = "Fruits", Name = "Apple" },
                new Item { Category = "Fruits", Name = "Banana" },
                new Item { Category = "Fruits", Name = "Cherry" },
                new Item { Category = "Vegetables", Name = "Carrot" },
                new Item { Category = "Vegetables", Name = "Lettuce" },
                new Item { Category = "Grains", Name = "Rice" },
                new Item { Category = "Grains", Name = "Wheat" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write the LINQ Reporting tags.
        // Outer foreach iterates over groups created by GroupBy on Category.
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");
        // Group title (category name).
        builder.Writeln("<<[g.Key]>>");

        // Apply a bulleted list to the items of each group.
        builder.ListFormat.List = template.Lists.Add(ListTemplate.BulletDefault);
        builder.Writeln("<<foreach [item in g]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        // End the bullet list for this group.
        builder.ListFormat.RemoveNumbers();

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
}
