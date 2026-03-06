// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

namespace AsposeWordsLinqReportingPrint
{
    // Simple POCO that will be used as the data source for the report.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    // Wrapper class that holds the collection of products.
    public class ReportData
    {
        public List<Product> Products { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOCM template that contains Aspose.Words reporting tags.
            //    The Document(string) constructor automatically detects the format.
            Document doc = new Document(@"C:\Templates\ReportTemplate.docm");

            // 2. Prepare the data source.
            var data = new ReportData
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 0.99m },
                    new Product { Name = "Banana", Price = 0.59m },
                    new Product { Name = "Cherry", Price = 2.49m }
                }
            };

            // 3. Populate the template using the ReportingEngine.
            //    The data source is referenced in the template by the name "ds".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "ds");

            // 4. Print the resulting document.
            //    This prints to the default printer without showing any UI.
            doc.Print();

            // Optional: if you need to specify printer settings (e.g., page range),
            // create a PrinterSettings object and use the overload that accepts it.
            // PrinterSettings settings = new PrinterSettings
            // {
            //     PrinterName = "Your Printer Name",
            //     PrintRange = PrintRange.SomePages,
            //     FromPage = 1,
            //     ToPage = 2
            // };
            // doc.Print(settings);
        }
    }
}
