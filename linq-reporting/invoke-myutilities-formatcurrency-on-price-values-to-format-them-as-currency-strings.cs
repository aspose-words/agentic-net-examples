using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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

// Data model for a product.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

// Wrapper model that holds the collection used in the template.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln(); // Empty line for readability.

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name : <<[p.Name]>>");
        // Invoke the custom static method to format the price.
        builder.Writeln("Price: <<[MyUtilities.FormatCurrency(p.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // Sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new() { Name = "Apple",  Price = 1.23m },
                new() { Name = "Banana", Price = 0.99m },
                new() { Name = "Cherry", Price = 2.50m }
            }
        };

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register the utility class so its static method can be called from the template.
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Build the report. No data source name is needed because the template references members directly.
        engine.BuildReport(doc, model);

        // Save the generated report.
        doc.Save(reportPath);
    }
}
