using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

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
                new Product { Name = "Coffee", Price = 4.99m }
            }
        };

        // Build the template document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Products]>>");
        // Use a pre‑formatted property to render the price as currency.
        builder.Writeln("<<[p.Name]>> - <<[p.FormattedPrice]>>");
        builder.Writeln("<</foreach>>");

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("ReportCurrency.docx");
    }
}

// Wrapper class that serves as the root data source for the report.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Simple product class with public properties.
// Includes a read‑only property that returns the price formatted as currency.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }

    // Returns the price formatted using the standard .NET currency format string.
    public string FormattedPrice => Price.ToString("c");
}
