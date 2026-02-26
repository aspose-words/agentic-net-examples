using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple POCO class that will be used as a data source.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains LINQ expressions.
            // The template can contain tags like <<foreach [in ds.Products]>><<[Name]>><</foreach>>.
            Document template = new Document("Template.docm");

            // Prepare a collection of products.
            List<Product> allProducts = new List<Product>
            {
                new Product { Name = "Apple",  Price = 5.0 },
                new Product { Name = "Banana", Price = 2.5 },
                new Product { Name = "Laptop", Price = 1200.0 },
                new Product { Name = "Desk",   Price = 250.0 }
            };

            // Use LINQ to filter the collection – only expensive items (price > 100) will be shown.
            // The result of the LINQ query is stored in an anonymous object that will be passed to the engine.
            var dataSource = new
            {
                // The property name "Products" will be used in the template.
                Products = allProducts.Where(p => p.Price > 100).ToList()
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument ("ds") is the name used in the template to reference the data source.
            // Example template tag: <<foreach [in ds.Products]>><<[Name]>><</foreach>>.
            engine.BuildReport(template, dataSource, "ds");

            // Save the populated document.
            template.Save("Report.docx");
        }
    }
}
