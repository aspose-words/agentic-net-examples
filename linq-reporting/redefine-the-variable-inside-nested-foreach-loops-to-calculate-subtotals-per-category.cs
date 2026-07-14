using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // 2. Prepare the data model.
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

        // 3. Load the template and build the report.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 4. Save the generated report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }

    // Creates a simple Word document containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Category Report");
        builder.Writeln();

        // Outer foreach: iterate over categories.
        builder.Writeln("<<foreach [cat in model.Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln();

        // Inner foreach: iterate over items within a category.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Subtotal for the current category.
        builder.Writeln("Subtotal: $<<[cat.Subtotal]>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Root data model passed to the reporting engine.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Represents a category containing multiple items.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();

    // Calculated subtotal for the category.
    public decimal Subtotal => CalculateSubtotal();

    private decimal CalculateSubtotal()
    {
        decimal sum = 0;
        foreach (var i in Items)
            sum += i.Price;
        return sum;
    }
}

// Represents an individual item.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
