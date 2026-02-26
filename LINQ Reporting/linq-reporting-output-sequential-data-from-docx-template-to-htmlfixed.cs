using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data model that will be used as the LINQ data source.
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
            // Path to the DOCX template that contains LINQ Reporting tags.
            // Example tag in the template: <<foreach [ds.Products]>><<[Name]>> - <<[Price]:currency>> (Qty: <<[Quantity]>>)<</foreach>>
            string templatePath = @"C:\Templates\ProductsReport.docx";

            // Path where the resulting HTML Fixed document will be saved.
            string outputPath = @"C:\Output\ProductsReport.html";

            // Create a list of products – this will be the source collection for the LINQ query.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.60m, Quantity = 120 },
                new Product { Name = "Banana", Price = 0.40m, Quantity = 200 },
                new Product { Name = "Orange", Price = 0.80m, Quantity = 150 }
            };

            // Build an anonymous object that contains the collection.
            // The property name ("Products") must match the name used in the template.
            var dataSource = new
            {
                Products = products
            };

            // Load the template document.
            Document doc = new Document(templatePath);

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third argument ("ds") is the name by which the template will reference the data source object.
            engine.BuildReport(doc, dataSource, "ds");

            // Prepare save options for HTML Fixed format.
            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                SaveFormat = SaveFormat.HtmlFixed,
                // Optional: embed images as separate files rather than Base64.
                ExportEmbeddedImages = false,
                // Optional: specify a folder for external resources (images, CSS, etc.).
                ResourcesFolder = @"C:\Output\Resources",
                // Optional: do not show page borders in the generated HTML.
                ShowPageBorder = false
            };

            // Ensure the resources folder exists.
            System.IO.Directory.CreateDirectory(saveOptions.ResourcesFolder);

            // Save the populated document as HTML Fixed.
            doc.Save(outputPath, saveOptions);
        }
    }
}
