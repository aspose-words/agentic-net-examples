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

    class Program
    {
        static void Main()
        {
            // Load the template document that contains a table with a foreach tag:
            // <<foreach [in products]>><<[Name]>> <<[Price]>> <<[Quantity]>> <</foreach>>
            Document template = new Document("Template.docx");

            // Prepare a collection of products to be merged into the template.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.50m, Quantity = 120 },
                new Product { Name = "Banana", Price = 0.30m, Quantity = 200 },
                new Product { Name = "Orange", Price = 0.80m, Quantity = 150 }
            };

            // The ReportingEngine will iterate over the "products" collection defined in the template.
            // We pass an anonymous object that exposes the collection as a property named "products".
            ReportingEngine engine = new ReportingEngine();

            // BuildReport populates the template using the provided data source.
            // The second parameter is the data source object; the third (optional) parameter can be omitted
            // because we reference the collection directly by its property name.
            engine.BuildReport(template, new { products = products });

            // Save the resulting document.
            template.Save("Report.docx");
        }
    }
}
