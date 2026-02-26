using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model that will be used as the data source for the report.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public bool InStock { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a list of products that will be merged into the template.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Laptop", Price = 1299.99m, InStock = true },
                new Product { Name = "Mouse",   Price =  25.50m, InStock = true },
                new Product { Name = "Desk",    Price = 300.00m, InStock = false }
            };

            // Load the MHTML template that contains the Reporting Engine syntax.
            // The template should include a conditional block, e.g.:
            // <<if [product.InStock]>><<[product.Name]>> - <<[product.Price]:currency>><</if>>
            Document doc = new Document("Template.mht");

            // Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after the conditional block is evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the list of products as the data source.
            // The data source name ("product") must match the name used in the template.
            engine.BuildReport(doc, products, "product");

            // Save the populated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
