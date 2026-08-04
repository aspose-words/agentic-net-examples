using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Helper class whose static members will be used in the template.
    public static class Formatter
    {
        // Formats a price value as a currency string.
        public static string FormatPrice(double price) => $"${price:F2}";
    }

    // Data model exposed to the template.
    public class ReportModel
    {
        // Collection of products to iterate over in the template.
        public List<Product> Products { get; set; } = new();
    }

    // Simple product entity.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin a foreach loop over the Products collection.
            builder.Writeln("<<foreach [p in Products]>>");
            // Output product name and formatted price using the static helper.
            builder.Writeln("<<[p.Name]>> - <<[Formatter.FormatPrice(p.Price)]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back before building the report.
            // -----------------------------------------------------------------
            var document = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare a large data set.
            // -----------------------------------------------------------------
            var model = new ReportModel();
            for (int i = 1; i <= 1000; i++)
            {
                model.Products.Add(new Product
                {
                    Name = $"Product {i}",
                    Price = i * 1.23 // Example price.
                });
            }

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            // Enable reflection optimization for maximum performance.
            ReportingEngine.UseReflectionOptimization = true;

            var engine = new ReportingEngine();

            // Register the external type so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(Formatter));

            // Build the report. The root object name must match the name used in the template tags.
            engine.BuildReport(document, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            document.Save(outputPath);
        }
    }
}
