using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Formats a numeric value as a currency string using the current culture.
    public static string FormatCurrency(decimal value) => value.ToString("C", CultureInfo.CurrentCulture);
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare the template document.
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in model.Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Price: <<[MyUtilities.FormatCurrency(p.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk and reload it to satisfy the lifecycle rule.
        doc.Save(templatePath);
        var templateDoc = new Document(templatePath);

        // Create sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new() { Name = "Apple", Price = 1.23m },
                new() { Name = "Banana", Price = 0.99m },
                new() { Name = "Cherry", Price = 2.50m }
            }
        };

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        templateDoc.Save(outputPath);
    }
}
