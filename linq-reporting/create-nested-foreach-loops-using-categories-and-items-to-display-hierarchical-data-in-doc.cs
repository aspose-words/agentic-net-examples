using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Categories = new()
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new()
                    {
                        new Item { Name = "Apple", Price = 0.5m },
                        new Item { Name = "Banana", Price = 0.3m }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new()
                    {
                        new Item { Name = "Carrot", Price = 0.2m },
                        new Item { Name = "Tomato", Price = 0.4m }
                    }
                }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.Writeln("- <<[itm.Name]>>: $<<[itm.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
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
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}
