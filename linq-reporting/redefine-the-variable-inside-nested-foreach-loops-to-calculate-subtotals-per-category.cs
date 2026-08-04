using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
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
                        new Item { Name = "Broccoli", Price = 1.10m }
                    }
                }
            }
        };

        // Calculate subtotals per category (redefining the variable inside nested loops conceptually).
        foreach (var cat in model.Categories)
        {
            decimal subtotal = 0;
            foreach (var itm in cat.Items)
                subtotal += itm.Price;
            cat.Subtotal = subtotal;
        }

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Sales Report");
        builder.Writeln();

        // Outer foreach over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln();

        // Inner foreach over items.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Subtotal for the current category.
        builder.Writeln("Subtotal: $<<[cat.Subtotal]>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
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
