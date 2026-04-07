using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

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
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Create a bulleted list style.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList; // Apply the list to following paragraphs.

        // Outer foreach – groups by Category.
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");

        // Group header (first level bullet).
        builder.Writeln("<<[g.Key]>>");

        // Increase list level for inner items.
        builder.ListFormat.ListIndent();

        // Inner foreach – items within the current group.
        builder.Writeln("<<foreach [item in g]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Return to the outer list level.
        builder.ListFormat.ListOutdent();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var document = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the data source.
        engine.BuildReport(document, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        document.Save(outputPath);
    }
}
