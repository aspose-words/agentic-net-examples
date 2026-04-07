using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Wrapper class that will be passed to the reporting engine.
    public class ReportModel
    {
        // Collection that the template will iterate over.
        public List<Product> Products { get; set; } = new();
    }

    // Simple data entity.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a DOCX template programmatically.
            // -----------------------------------------------------------------
            string templatePath = "Template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("Product List:");
            builder.Writeln();

            // Opening tag for the collection.
            builder.Writeln("<<foreach [product in Products]>>");

            // Content for each item.
            builder.Writeln("- <<[product.Name]>> : $<<[product.Price]>>");

            // Closing tag for the collection.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 0.99 },
                    new Product { Name = "Banana", Price = 0.59 },
                    new Product { Name = "Cherry", Price = 2.49 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed.

            // Bind the data model to the template using the root name "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
