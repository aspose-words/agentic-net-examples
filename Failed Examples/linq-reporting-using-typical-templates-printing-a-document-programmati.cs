// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

namespace AsposeWordsReportingDemo
{
    // Simple data model that will be used as a LINQ data source.
    public class Product
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }

        public decimal Total => Quantity * Price;
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare a collection of products using LINQ (or any other query).
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Quantity = 10, Price = 0.5m },
                new Product { Name = "Banana", Quantity = 5,  Price = 0.3m },
                new Product { Name = "Orange", Quantity = 8,  Price = 0.4m }
            };

            // 2. Load the template document that contains Aspose.Words reporting tags.
            //    The template should have tags like <<foreach [products]>><<[Name]>> - <<[Quantity]>> - <<[Price]:currency>><</foreach>>
            string templatePath = @"C:\Docs\Template.docx";
            Document template = new Document(templatePath); // load rule

            // 3. Populate the template with the data source using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name by which the data source will be referenced inside the template.
            engine.BuildReport(template, products, "products");

            // 4. Save the generated report.
            string reportPath = @"C:\Docs\Report.docx";
            template.Save(reportPath); // save rule

            // 5. Print the generated report programmatically (no UI).
            //    Configure printer settings as needed.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example: print to the default printer, all pages.
                PrinterName = new PrinterSettings().PrinterName,
                PrintRange = PrintRange.AllPages
            };

            // Wrap the document in AsposeWordsPrintDocument which integrates with .NET printing.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(template);
            printDoc.PrinterSettings = printerSettings;

            // Optional: cache printer settings to speed up the first print call.
            printDoc.CachePrinterSettings();

            // Execute the print job.
            printDoc.Print();

            // 6. (Optional) Show a print preview dialog with the same document.
            //    Uncomment the following lines if a UI environment is available.
            /*
            System.Windows.Forms.PrintPreviewDialog previewDlg = new System.Windows.Forms.PrintPreviewDialog
            {
                Document = printDoc
            };
            previewDlg.ShowDialog();
            */
        }
    }
}
