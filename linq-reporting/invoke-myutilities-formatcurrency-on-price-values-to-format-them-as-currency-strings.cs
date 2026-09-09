using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Formats a numeric value as a currency string using the current culture.
    public static string FormatCurrency(double value)
    {
        return value.ToString("C");
    }
}

// Data model for a product.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}

// Wrapper model that contains a collection of products.
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
                new Product { Name = "Apple",  Price = 1.23 },
                new Product { Name = "Banana", Price = 0.99 },
                new Product { Name = "Cherry", Price = 2.50 }
            }
        };

        // Create a new blank document and a builder to construct the template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a title.
        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");
        // Output product name.
        builder.Writeln("Name : <<[p.Name]>>");
        // Output formatted price using the custom utility method.
        builder.Writeln("Price: <<[MyUtilities.FormatCurrency(p.Price)]>>");
        builder.Writeln(); // Add an empty line between items.
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register the utility class so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Build the report using the model as the root data source.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
