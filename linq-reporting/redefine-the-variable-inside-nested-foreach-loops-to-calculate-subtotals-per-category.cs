using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Beverages",
                    Items = new List<Item>
                    {
                        new Item { Name = "Coffee", Price = 3.50m },
                        new Item { Name = "Tea", Price = 2.75m },
                        new Item { Name = "Juice", Price = 4.20m }
                    }
                },
                new Category
                {
                    Name = "Snacks",
                    Items = new List<Item>
                    {
                        new Item { Name = "Chips", Price = 1.99m },
                        new Item { Name = "Cookies", Price = 2.49m }
                    }
                }
            }
        };

        // Create the template document programmatically.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("- <<[item.Name]>>: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Subtotal for <<[cat.Name]>>: $<<[cat.Items.Sum(i => i.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(templateDoc, model, "model");
        if (!success)
        {
            Console.WriteLine("Report generation failed.");
            return;
        }

        // Save the generated report.
        const string outputPath = "report.docx";
        templateDoc.Save(outputPath);
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
