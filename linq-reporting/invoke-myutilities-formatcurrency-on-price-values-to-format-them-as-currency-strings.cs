using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Formats a decimal value as a currency string using the current culture.
    public static string FormatCurrency(decimal value)
    {
        return value.ToString("C", CultureInfo.CurrentCulture);
    }
}

// Wrapper that exposes the utility method to the reporting engine.
public class MyUtilitiesHelper
{
    public string FormatCurrency(decimal value) => MyUtilities.FormatCurrency(value);
}

// Data model classes.
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Create the template document programmatically.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // Use the utility wrapper to format the price.
        builder.Writeln("Price: <<[MyUtilities.FormatCurrency(p.Price)]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Prepare the utility wrapper instance.
        var utils = new MyUtilitiesHelper();

        // Build the report using two data sources: the model and the utilities wrapper.
        var engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { model, utils }, new[] { "model", "MyUtilities" });

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
