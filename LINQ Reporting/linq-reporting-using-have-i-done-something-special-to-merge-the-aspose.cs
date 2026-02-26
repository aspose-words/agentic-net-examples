using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data entity.
    public class Product
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a list of products – this will be the source for the LINQ query.
            List<Product> allProducts = new List<Product>
            {
                new Product { Name = "Apple iPhone 15", Category = "Electronics", Price = 999.99m },
                new Product { Name = "Samsung Galaxy S24", Category = "Electronics", Price = 899.99m },
                new Product { Name = "Dell XPS 13", Category = "Computers", Price = 1199.00m },
                new Product { Name = "HP Envy", Category = "Computers", Price = 999.00m },
                new Product { Name = "Sony WH‑1000XM5", Category = "Audio", Price = 349.99m }
            };

            // 2. Use LINQ to select only the products we want to appear in the report.
            //    For example, all electronics cheaper than $950.
            var filteredProducts = allProducts
                .Where(p => p.Category == "Electronics" && p.Price < 950m)
                .Select(p => new
                {
                    // The anonymous type's property names must match the merge fields in the template.
                    ProductName = p.Name,
                    ProductPrice = p.Price
                })
                .ToList();

            // 3. Load the Word template that contains the merge fields.
            //    The template should have a table with a repeatable region:
            //    <<foreach [ds]>>
            //        <<[ProductName]>>
            //        <<[ProductPrice]:currency>>
            //    <</foreach>>
            Document doc = new Document("Template.docx");   // lifecycle rule: load

            // 4. Create the ReportingEngine and populate the template.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" must match the name used in the template tags.
            engine.BuildReport(doc, filteredProducts, "ds"); // feature rule: use BuildReport

            // 5. Save the generated report.
            doc.Save("FilteredProductsReport.docx");        // lifecycle rule: save
        }
    }
}
