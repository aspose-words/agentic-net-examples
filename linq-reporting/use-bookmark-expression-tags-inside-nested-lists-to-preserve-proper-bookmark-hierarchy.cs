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
                        new Item { Name = "Apple" },
                        new Item { Name = "Banana" }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new List<Item>
                    {
                        new Item { Name = "Carrot" },
                        new Item { Name = "Tomato" }
                    }
                }
            }
        };

        // Create a template document with nested lists and bookmark tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with nested lists and bookmarks:");
        builder.Writeln();

        // Outer list: categories.
        builder.Writeln("<<foreach [category in Categories]>>");
        builder.Writeln("- <<bookmark [category.Name]>><<[category.Name]>> <</bookmark>>");
        // Inner list: items within each category.
        builder.Writeln("  <<foreach [item in category.Items]>>");
        builder.Writeln("  * <<bookmark [item.Name]>><<[item.Name]>> <</bookmark>>");
        builder.Writeln("  <</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine("Report generated successfully at: " + reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Category with a collection of items.
public class Category
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

// Simple item model.
public class Item
{
    public string Name { get; set; } = "";
}
