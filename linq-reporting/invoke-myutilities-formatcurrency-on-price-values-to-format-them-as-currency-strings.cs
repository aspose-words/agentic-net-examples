using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.23m },
                new Product { Name = "Banana", Price = 0.99m },
                new Product { Name = "Cherry", Price = 2.50m }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("<<[p.Name]>> - <<[MyUtilities.FormatCurrency(p.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before BuildReport).
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Register the utility class so its static method can be used in the template.
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(loadedDoc, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);

        // Inform that the process completed (no interactive input required).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

// ---------------------------------------------------------------------
// Utility class containing the currency formatting method.
// ---------------------------------------------------------------------
public static class MyUtilities
{
    // Formats a decimal value as a currency string using the current culture.
    public static string FormatCurrency(decimal value)
    {
        return value.ToString("C", CultureInfo.CurrentCulture);
    }
}
