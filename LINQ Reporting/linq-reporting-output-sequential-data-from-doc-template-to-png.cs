using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data class that will be used as the data source for the report.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains LINQ Reporting tags, e.g. <<foreach [product]>><<[Name]>> - <[Price]>><</foreach>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a list of products as the data source.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m },
                new Product { Name = "Banana", Price = 0.80m },
                new Product { Name = "Cherry", Price = 2.50m }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("product") must match the name used in the template tags.
            engine.BuildReport(doc, products, "product");

            // Render each page of the populated document to a separate PNG image.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Configure image save options for PNG format.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    // Render only the current page.
                    PageSet = new PageSet(pageIndex)
                };

                // Save the page as PNG. The file name includes the page number (1‑based).
                string outputPath = $@"C:\Output\Report_Page_{pageIndex + 1}.png";
                doc.Save(outputPath, options);
            }
        }
    }
}
