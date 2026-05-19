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
            Categories = new List<Category>
            {
                new()
                {
                    Name = "Fruits",
                    Items = new List<Item>
                    {
                        new() { Name = "Apple", Price = 1.20m },
                        new() { Name = "Banana", Price = 0.80m },
                        new() { Name = "Orange", Price = 1.50m }
                    }
                },
                new()
                {
                    Name = "Vegetables",
                    Items = new List<Item>
                    {
                        new() { Name = "Carrot", Price = 0.60m },
                        new() { Name = "Broccoli", Price = 1.10m }
                    }
                }
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("- <<[item.Name]>>: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Subtotal: $<<[cat.Subtotal]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();

    // Subtotal calculated per category.
    public decimal Subtotal => Items.Sum(i => i.Price);
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}
