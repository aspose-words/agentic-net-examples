using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingCurrencyExample
{
    // Utility class for formatting currency.
    public static class MyUtilities
    {
        public static string FormatCurrency(decimal value)
        {
            // Format using the current culture's currency format.
            return value.ToString("C", CultureInfo.CurrentCulture);
        }
    }

    // Data model for a product.
    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }

        // Returns the price formatted as currency using MyUtilities.
        public string FormattedPrice => MyUtilities.FormatCurrency(Price);
    }

    // Root model passed to the reporting engine.
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
                    new Product { Name = "Apple", Price = 1.25m },
                    new Product { Name = "Banana", Price = 0.75m },
                    new Product { Name = "Cherry", Price = 2.50m }
                }
            };

            // Create a temporary folder for files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create the template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write a title.
            builder.Writeln("Product Price Report");
            builder.Writeln();

            // Begin foreach tag for products.
            builder.Writeln("<<foreach [p in Products]>>");

            // Create a table with two columns: Name and Formatted Price.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            // Use the formatted price property.
            builder.Writeln("<<[p.FormattedPrice]>>");
            builder.EndRow();

            builder.EndTable();

            // End foreach tag.
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var doc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);
        }
    }
}
