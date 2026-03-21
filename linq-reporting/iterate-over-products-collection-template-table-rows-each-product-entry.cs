using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsTableExample
{
    // Simple data model representing a product.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    // Wrapper class required by ReportingEngine.
    public class ReportData
    {
        // Title used in the template (optional, but prevents missing‑field errors).
        public string Title { get; set; }

        // Collection of products to iterate over.
        public List<Product> products { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a template document in memory with the required tags.
            Document template = new Document();
            var builder = new DocumentBuilder(template);

            // Optional title placeholder.
            builder.Writeln("<<[Title]>>");
            builder.Writeln();

            // Table header (optional, for readability).
            builder.Writeln("Name\tPrice\tQuantity");
            builder.Writeln();

            // Foreach tag that will be replaced by the product rows.
            builder.Writeln("<<foreach [in products]>><<[Name]>>\t<<[Price]>>\t<<[Quantity]>> <</foreach>>");

            // Prepare a collection of products to be merged into the template.
            var products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.50m, Quantity = 120 },
                new Product { Name = "Banana", Price = 0.30m, Quantity = 200 },
                new Product { Name = "Orange", Price = 0.80m, Quantity = 150 }
            };

            // Create the data source wrapper.
            var data = new ReportData
            {
                Title = "Product Report",
                products = products
            };

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(template, data);

            // Save the resulting document.
            template.Save("Report.docx");
        }
    }
}
