using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Simple data model for the report
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Category { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains a multicolored numbered list.
            // Example template content (using ReportingEngine syntax):
            //   <<foreach [ds.Products]>>
            //   <<[Name]>>
            //   <<[Price]:currency>>
            //   <<[Category]>>
            //   <</foreach>>
            string templatePath = @"C:\Templates\MulticoloredNumberedListTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a data source – a list of products.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",   Price = 0.99m, Category = "Fruit" },
                new Product { Name = "Banana",  Price = 0.59m, Category = "Fruit" },
                new Product { Name = "Carrot",  Price = 0.39m, Category = "Vegetable" },
                new Product { Name = "Detergent", Price = 3.49m, Category = "Cleaning" }
            };

            // Use LINQ to order the products by category and then by price.
            var orderedProducts = products
                .OrderBy(p => p.Category)
                .ThenBy(p => p.Price)
                .ToList();

            // Wrap the ordered list in an anonymous object so that the template can reference it as "Products".
            var dataSource = new { Products = orderedProducts };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The second overload allows the template to reference the data source object itself via the name "ds".
            engine.BuildReport(doc, dataSource, "ds");

            // Save the populated document.
            string outputPath = @"C:\Reports\LinqMulticoloredListReport.docx";
            doc.Save(outputPath);
        }
    }
}
