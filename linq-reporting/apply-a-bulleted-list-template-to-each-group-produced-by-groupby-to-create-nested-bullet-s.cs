using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;          // Needed for ListTemplate
using Aspose.Words.Reporting;

public class Product
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
}

public class CategoryGroup
{
    public string Category { get; set; } = "";
    public List<Product> Items { get; set; } = new();
}

public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var products = new List<Product>
        {
            new() { Category = "Fruits", Name = "Apple" },
            new() { Category = "Fruits", Name = "Banana" },
            new() { Category = "Fruits", Name = "Cherry" },
            new() { Category = "Vegetables", Name = "Carrot" },
            new() { Category = "Vegetables", Name = "Lettuce" },
            new() { Category = "Beverages", Name = "Coffee" },
            new() { Category = "Beverages", Name = "Tea" }
        };

        // Group by category.
        var model = new ReportModel
        {
            Groups = products
                .GroupBy(p => p.Category)
                .Select(g => new CategoryGroup
                {
                    Category = g.Key,
                    Items = g.ToList()
                })
                .ToList()
        };

        // Create template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Apply a bulleted list to the outer level.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);

        // Outer foreach – groups.
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("<<[g.Category]>>");

        // Increase list level for inner items.
        builder.ListFormat.ListIndent();

        // Inner foreach – products.
        builder.Writeln("<<foreach [p in g.Items]>>");
        builder.Writeln("<<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Return to outer list level.
        builder.ListFormat.ListOutdent();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save template (optional).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, model, "model");

        // Save final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
