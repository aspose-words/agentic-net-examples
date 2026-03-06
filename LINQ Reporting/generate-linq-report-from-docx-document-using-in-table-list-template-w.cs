using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains an in‑table list template with alternate content.
        Document template = new Document("Template.docx");

        // Create a collection of products using LINQ.
        List<Product> products = new List<Product>
        {
            new Product { Name = "Apple",  Category = "Fruit",      Price = 1.20 },
            new Product { Name = "Banana", Category = "Fruit",      Price = 1.10 },
            new Product { Name = "Carrot", Category = "Vegetable", Price = 0.80 },
            new Product { Name = "Lettuce",Category = "Vegetable", Price = 0.90 }
        };

        // Group products by category – this will be used as the master list in the template.
        // Each group contains a list of items that the in‑table list will iterate over.
        List<CategoryGroup> data = products
            .GroupBy(p => p.Category)
            .Select(g => new CategoryGroup
            {
                Category = g.Key,
                Items = g.ToList()
            })
            .ToList();

        // Build the report using Aspose.Words ReportingEngine.
        // The data source name ("Categories") must match the name used in the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "Categories");

        // Save the populated document.
        template.Save("Report.docx");
    }

    // Simple POCO representing a product.
    public class Product
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public double Price { get; set; }
    }

    // Wrapper class for a category group – used for the master‑detail structure.
    public class CategoryGroup
    {
        public string Category { get; set; }
        public List<Product> Items { get; set; }
    }
}
