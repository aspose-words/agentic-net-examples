using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
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
            // Load the MHTML template that contains the reporting syntax.
            // The template should have a table with a conditional row, e.g.:
            // <<foreach [Products]>>
            //   <<[Name]>>  <<[Price]>>  <<if [InStock]>>In stock<<else>>Out of stock<<endif>>
            // <<endforeach>>
            Document template = new Document("Template.mht");

            // Prepare the data source – a list of products.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m, InStock = true  },
                new Product { Name = "Banana", Price = 0.80m, InStock = false },
                new Product { Name = "Cherry", Price = 2.50m, InStock = true  }
            };

            // Wrap the list in an anonymous object so the template can reference it as "Products".
            var dataSource = new { Products = products };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Optional: remove empty paragraphs that may appear after conditional blocks are evaluated.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report – populate the template with the data source.
            // The third parameter (dataSourceName) is left empty because we reference members directly.
            engine.BuildReport(template, dataSource, "");

            // Save the generated document.
            template.Save("Report.docx");
        }
    }
}
