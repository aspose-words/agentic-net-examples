using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample hierarchical data.
        ReportModel model = new()
        {
            Categories = new()
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new()
                    {
                        new Item { Name = "Apple" },
                        new Item { Name = "Banana" },
                        new Item { Name = "Cherry" }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new()
                    {
                        new Item { Name = "Carrot" },
                        new Item { Name = "Lettuce" },
                        new Item { Name = "Tomato" }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer foreach over categories.
        builder.Writeln("<<foreach [category in Categories]>>");
        builder.Writeln("Category: <<[category.Name]>>");

        // Inner foreach over items of the current category.
        builder.Writeln("<<foreach [item in category.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>"); // End inner foreach.

        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object name used in the template is "model".
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, with public properties, non‑nullable).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
}
