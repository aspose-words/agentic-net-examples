using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data model representing a product.
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
            // Prepare sample data.
            var products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m, InStock = true  },
                new Product { Name = "Banana", Price = 0.80m, InStock = false },
                new Product { Name = "Cherry", Price = 2.50m, InStock = true  },
                new Product { Name = "Date",   Price = 3.00m, InStock = false }
            };

            // Load the Word template that contains LINQ Reporting tags.
            // The template should define a table with a conditional block, e.g.:
            // <<foreach [Products]>>
            //   <<if [InStock]>>
            //     <<[Name]>>    <<[Price]>>    (In stock)
            //   <<else>>
            //     <<[Name]>>    <<[Price]>>    (Out of stock)
            //   <<endif>>
            // <<endforeach>>
            Document template = new Document("Template.docx");

            // Configure the reporting engine.
            var engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional blocks are evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using an anonymous object as the data source.
            // The property name "Products" matches the tag used in the template.
            engine.BuildReport(template, new { Products = products });

            // Save the populated document as PDF.
            var pdfOptions = new PdfSaveOptions();
            template.Save("Report.pdf", pdfOptions);
        }
    }
}
