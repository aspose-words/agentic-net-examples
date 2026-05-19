using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Formats a numeric value as a currency string using the current culture.
    public static string FormatCurrency(decimal value)
    {
        return value.ToString("C", CultureInfo.CurrentCulture);
    }
}

// Data model classes.
public class Product
{
    public string Name { get; set; } = "";
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.25m },
                new Product { Name = "Banana", Price = 0.75m },
                new Product { Name = "Cherry", Price = 2.50m }
            }
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Products]>>");
        // Call the static utility method to format the price.
        builder.Writeln("<<[MyUtilities.FormatCurrency(p.Price)]>> - <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();

        // Register the static class so its methods can be used in the template.
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Build the report using the root object name "model".
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
