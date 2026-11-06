using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Categories = new()
            {
                new Category
                {
                    Name = "Beverages",
                    Items = new()
                    {
                        new Item { Name = "Coffee", Amount = 3.50m },
                        new Item { Name = "Tea", Amount = 2.75m },
                        new Item { Name = "Juice", Amount = 4.20m }
                    }
                },
                new Category
                {
                    Name = "Snacks",
                    Items = new()
                    {
                        new Item { Name = "Chips", Amount = 1.99m },
                        new Item { Name = "Cookies", Amount = 2.49m }
                    }
                }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("LINQ Reporting – Subtotals per Category");
        builder.Writeln();

        // Outer foreach over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln();

        // Inner foreach over items.
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.Writeln("- <<[itm.Name]>>: $<<[itm.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Subtotal for the current category.
        builder.Writeln("Subtotal: $<<[cat.Subtotal]>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root model.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Category with a computed subtotal.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();

    // Subtotal calculated from the items.
    public decimal Subtotal => Items.Sum(i => i.Amount);
}

// Individual item.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}
