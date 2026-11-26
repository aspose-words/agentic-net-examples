using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Sample data model.
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
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Product List:");
            builder.Writeln("<<foreach [p in Products]>>");
            builder.Writeln("- <<[p.Name]>>: $<<[p.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new(templatePath);

            // Prepare sample data.
            ReportModel model = new()
            {
                Products = new()
                {
                    new Product { Name = "Apple", Price = 0.99m },
                    new Product { Name = "Banana", Price = 0.59m },
                    new Product { Name = "Cherry", Price = 2.49m }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
