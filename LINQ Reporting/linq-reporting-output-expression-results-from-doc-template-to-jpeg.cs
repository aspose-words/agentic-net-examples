using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data model used for the LINQ query.
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
            // Path to the Word template that contains LINQ Reporting tags.
            // Example tag in the template:
            // <<foreach [in ds]>><<[Name]>> - <<[Total]>>\n<</foreach>>
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare source data.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.5m, Quantity = 10 },
                new Product { Name = "Banana", Price = 0.3m, Quantity = 15 },
                new Product { Name = "Orange", Price = 0.8m, Quantity = 7 }
            };

            // LINQ query that calculates the total price for each product.
            var reportData = products
                .Select(p => new
                {
                    Name = p.Name,
                    // Expression result that will be used in the template.
                    Total = p.Price * p.Quantity
                })
                .ToList();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" must match the name used in the template tags.
            engine.BuildReport(doc, reportData, "ds");

            // Save the populated document as a JPEG image (first page only).
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page; change PageSet to render other pages.
                PageSet = new PageSet(0)
            };

            string outputPath = @"C:\Docs\Report.jpg";
            doc.Save(outputPath, jpegOptions);
        }
    }
}
