using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    public static string FormatCurrency(decimal price)
    {
        return price.ToString("C", CultureInfo.CurrentCulture);
    }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }

    // Helper property that formats the price using the utility method.
    public string FormattedPrice => MyUtilities.FormatCurrency(Price);
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>> - Price: <<[p.FormattedPrice]>>");
        builder.Writeln("<</foreach>>");

        // Ensure output directories exist.
        Directory.CreateDirectory("Output");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Output/Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
