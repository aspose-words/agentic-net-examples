using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for maximum performance.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a simple foreach loop that uses the registered external type.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Name]>> - <<[ExternalFormatter.Format(item.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template back (required before building the report).
        var loadedTemplate = new Document(templatePath);

        // Prepare sample data with a large collection.
        var model = new ReportModel
        {
            Items = new()
        };
        for (int i = 1; i <= 1000; i++)
        {
            model.Items.Add(new Item
            {
                Name = $"Product {i}",
                Price = i * 1.23m
            });
        }

        // Create the reporting engine and register the external type.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(ExternalFormatter));

        // Build the report.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}

// External type to be registered and used in the template.
public static class ExternalFormatter
{
    public static string Format(decimal price) => $"${price:F2}";
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item class used in the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
