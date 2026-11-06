using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCurrencyExample
{
    // Data model classes
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
            // Register code page provider (required by Aspose.Words for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple", Price = 0.99m },
                    new Product { Name = "Banana", Price = 0.59m },
                    new Product { Name = "Cherry", Price = 2.49m }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();

            // Document title
            builder.Writeln("Product Price Report");
            builder.Writeln();

            // Begin foreach loop over Products
            builder.Writeln("<<foreach [p in Products]>>");

            // Table for each product
            var table = builder.StartTable();

            // Header row
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Data row
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            // Render price as currency using standard .NET format string
            builder.Writeln("<<[p.Price.ToString(\"C\")]>>");
            builder.EndRow();

            // End table for this iteration
            builder.EndTable();

            // End foreach
            builder.Writeln("<</foreach>>");

            // Save the template
            builder.Document.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final report
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
