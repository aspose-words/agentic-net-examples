using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        Document template = new Document("Template.dotx");

        // Prepare a collection of products.
        var products = new List<Product>
        {
            new Product { Name = "Apple",  Price = 1.20 },
            new Product { Name = "Banana", Price = 0.80 },
            new Product { Name = "Cherry", Price = 2.50 }
        };

        // Use LINQ to filter and project the data that will be merged.
        var dataSource = products
            .Where(p => p.Price > 1.0)               // keep only expensive items
            .Select(p => new { p.Name, p.Price })    // create an anonymous type
            .ToArray();                               // convert to an array for the engine

        // Build the report by merging the LINQ data source into the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "Products");

        // Save the generated document.
        template.Save("Report.docx");
    }

    // Simple POCO class representing a product.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }
}
