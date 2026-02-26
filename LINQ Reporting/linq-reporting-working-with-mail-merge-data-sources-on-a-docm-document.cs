using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MailMergeLinqExample
{
    // Simple data entity that will be used as a LINQ data source.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains Reporting Engine tags, e.g. <<[p.Name]>> and <<[p.Price]:currency>>.
            Document template = new Document("Template.docm");

            // Create a LINQ data source – a list of Product objects.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            };

            // Initialise the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the template and the LINQ data source.
            // The third argument ("p") is the name by which the template refers to the data source.
            engine.BuildReport(template, products, "p");

            // Save the populated document.
            template.Save("ReportOutput.docx");
        }
    }
}
