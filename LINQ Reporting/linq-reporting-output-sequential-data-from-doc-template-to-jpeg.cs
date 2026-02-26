using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Sample data class that will be used as the data source for the LINQ Reporting Engine.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains LINQ Reporting tags, e.g. <<foreach [products]>><<[Name]>> - <<[Price]:currency>><</foreach>>
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source – a list of products.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            };

            // Build the report by populating the template with the data source.
            // The data source name ("products") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, products, "products");

            // Configure image save options to render the first page of the populated document as JPEG.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),

                // Optional: set JPEG quality (0‑100). Higher value = better quality, larger file.
                JpegQuality = 90
            };

            // Save the rendered page to a JPEG file.
            const string outputPath = @"C:\Output\ReportPage1.jpg";
            doc.Save(outputPath, saveOptions);
        }
    }
}
