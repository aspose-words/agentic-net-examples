using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data entity.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    // Extension method that will be used inside the template via LINQ syntax.
    public static class ProductExtensions
    {
        // Formats the price as currency.
        public static string PriceFormatted(this Product product)
        {
            return product.Price.ToString("C");
        }
    }

    // Model that will be passed to the ReportingEngine.
    public class ReportModel
    {
        public List<Product> Products { get; set; }

        public ReportModel()
        {
            Products = new List<Product>();
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains LINQ reporting tags.
            // Example tag in the template:
            // <<foreach [product in model.Products]>>
            //   <<[product.Name]>> – <<[product.PriceFormatted]>>
            // <<end>>
            Document doc = new Document("Template.docm");

            // Prepare the data source.
            var model = new ReportModel();
            model.Products.Add(new Product { Name = "Apple", Price = 1.20m });
            model.Products.Add(new Product { Name = "Banana", Price = 0.80m });
            model.Products.Add(new Product { Name = "Cherry", Price = 2.50m });

            // Build the report. The third argument is the name used inside the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the populated document.
            doc.Save("Report.docx");
        }
    }
}
