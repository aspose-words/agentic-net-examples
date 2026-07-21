using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer loop over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln();

        // Inner loop over items of the current category.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("- <<[item.Name]>>: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Subtotal for the current category (calculated in the data model).
        builder.Writeln("Subtotal: $<<[cat.Subtotal]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        ReportModel model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple", Price = 1.20m },
                        new Item { Name = "Banana", Price = 0.80m },
                        new Item { Name = "Orange", Price = 1.50m }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new List<Item>
                    {
                        new Item { Name = "Carrot", Price = 0.60m },
                        new Item { Name = "Broccoli", Price = 1.10m },
                        new Item { Name = "Spinach", Price = 0.90m }
                    }
                }
            }
        };

        // Compute subtotals for each category (redefining the variable per category).
        foreach (var cat in model.Categories)
        {
            decimal subtotal = 0m;
            foreach (var itm in cat.Items)
                subtotal += itm.Price;
            cat.Subtotal = subtotal;
        }

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
    public decimal Subtotal { get; set; }
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
