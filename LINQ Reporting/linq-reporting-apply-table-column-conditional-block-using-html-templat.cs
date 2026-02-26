using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Load the HTML template that contains LINQ Reporting syntax.
            // The template should be placed in the same folder as the executable or provide a full path.
            Document template = new Document("Template.html");

            // Prepare a data source that will be used by the ReportingEngine.
            // The anonymous object contains a collection named "Products".
            var dataSource = new
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",   Price = 1.20, ShowPrice = true  },
                    new Product { Name = "Banana",  Price = 0.80, ShowPrice = true  },
                    new Product { Name = "Cherry",  Price = 0.00, ShowPrice = false } // Price column will be hidden.
                }
            };

            // Create and configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after conditional blocks are evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The third argument ("data") is the name used inside the template
            // to reference the data source object itself (e.g., <<[data.Products]>>
            engine.BuildReport(template, dataSource, "data");

            // Save the populated document.
            template.Save("Report.docx");
        }

        // Simple POCO that represents a row in the table.
        public class Product
        {
            public string Name { get; set; }
            public double Price { get; set; }
            public bool ShowPrice { get; set; }
        }
    }
}
