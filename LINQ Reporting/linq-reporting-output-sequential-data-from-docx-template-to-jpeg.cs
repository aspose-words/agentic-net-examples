using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToJpeg
{
    // Sample data class that will be used as the data source for the LINQ reporting engine.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains LINQ reporting tags, e.g. <<foreach [ds.Products]>><<[Name]>> - <<[Price]>> - <<[Quantity]>> <</foreach>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document using the provided Document(string) constructor.
            Document doc = new Document(templatePath);

            // Prepare a list of products that will be bound to the template.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.50m, Quantity = 120 },
                new Product { Name = "Banana", Price = 0.30m, Quantity = 200 },
                new Product { Name = "Orange", Price = 0.80m, Quantity = 150 }
            };

            // Wrap the list in an anonymous object so that the template can reference it via the name "ds".
            var dataSource = new { Products = products };

            // Build the report using the LINQ ReportingEngine.
            // Use the overload that allows referencing the data source object itself in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "ds");

            // After the report is built, the document may span multiple pages.
            // Render each page to a separate JPEG image.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Adjust JPEG quality if needed (0‑100). 90 gives good quality with reasonable size.
                JpegQuality = 90
            };

            // Iterate through all pages of the populated document.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Set the PageSet to render only the current page (zero‑based index).
                jpegOptions.PageSet = new PageSet(pageIndex);

                // Construct the output file name, e.g. ReportPage_1.jpg, ReportPage_2.jpg, …
                string outputPath = $@"C:\Output\ReportPage_{pageIndex + 1}.jpg";

                // Save the current page as a JPEG using the provided Document.Save(string, SaveOptions) overload.
                doc.Save(outputPath, jpegOptions);
            }
        }
    }
}
