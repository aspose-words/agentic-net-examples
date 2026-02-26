// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDotmPrintExample
{
    // Simple data class for LINQ reporting
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains reporting tags (e.g., <<foreach [products]>><<[Name]>> - <<[Price]:currency>><</foreach>>)
            string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Load the DOTM template document
            Document doc = new Document(templatePath);

            // Prepare a LINQ data source
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            };

            // Build the report using Aspose.Words ReportingEngine
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("products") must match the name used in the template tags
            engine.BuildReport(doc, products, "products");

            // Ensure the page layout is up‑to‑date before printing
            doc.UpdatePageLayout();

            // Optional: configure printer settings (print to default printer in this example)
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Print all pages; modify as needed (e.g., FromPage, ToPage)
                PrintRange = PrintRange.AllPages
            };

            // Print the document using the specified printer settings
            doc.Print(printerSettings);
        }
    }
}
