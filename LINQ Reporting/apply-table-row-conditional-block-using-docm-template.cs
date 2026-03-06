using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsConditionalTable
{
    // Simple data model used by the template.
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
            // Load the DOCM template that contains a table with a conditional block.
            // The template should have tags like:
            //   <<foreach [product]>>
            //       <<if [product.InStock]>>
            //           <<[product.Name]>>
            //           <<[product.Price]:currency>>
            //       <<endif>>
            //   <<endforeach>>
            Document template = new Document(@"C:\Templates\ProductsTemplate.docm");

            // Prepare the data source – a list of products.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m, InStock = true  },
                new Product { Name = "Banana", Price = 0.80m, InStock = false },
                new Product { Name = "Carrot", Price = 0.50m, InStock = true  }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report: the engine will iterate over the list,
            // evaluate the conditional block (InStock) and populate the table rows.
            // The data source name "product" matches the tag used in the template.
            engine.BuildReport(template, products, "product");

            // Save the resulting document.
            template.Save(@"C:\Output\ProductsReport.docx", SaveFormat.Docx);
        }
    }
}
