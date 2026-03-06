using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity that will be used in the template.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    // Wrapper class that holds the collection for the template.
    public class DataSource
    {
        // The template will reference this property (e.g. <<foreach [ds.Products]>>).
        public List<Product> Products { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template that contains LINQ Reporting syntax.
            Document template = new Document("Template.dotx");

            // ----- Convert raw arrays to a canonical collection type (List<T>) -----
            // Example raw arrays.
            string[] productNames = { "Apple", "Banana", "Cherry" };
            double[] productPrices = { 1.20, 0.80, 2.50 };

            // Combine the parallel arrays into an array of Product objects.
            Product[] productArray = productNames
                .Select((name, index) => new Product
                {
                    Name = name,
                    Price = productPrices.Length > index ? productPrices[index] : 0.0
                })
                .ToArray();

            // Convert the array to a List<Product> – the canonical collection type expected by the engine.
            List<Product> productList = productArray.ToList();

            // Prepare the data source object that will be passed to the ReportingEngine.
            DataSource data = new DataSource
            {
                Products = productList
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // "ds" is the name used inside the template to reference the data source object.
            engine.BuildReport(template, data, "ds");

            // Save the populated document.
            template.Save("ReportOutput.docx");
        }
    }
}
