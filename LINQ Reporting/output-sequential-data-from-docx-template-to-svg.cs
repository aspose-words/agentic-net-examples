using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace DocxToSvgExample
{
    // Simple data source class used by the reporting engine.
    public class InvoiceData
    {
        public string InvoiceNumber { get; set; }
        public DateTime Date { get; set; }
        public string CustomerName { get; set; }
        public List<InvoiceItem> Items { get; set; }
    }

    public class InvoiceItem
    {
        public string Description { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal Total => Quantity * UnitPrice;
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains Aspose.Words reporting tags.
            Document doc = new Document("Template.docx");

            // Prepare sample data that matches the tags in the template.
            var data = new InvoiceData
            {
                InvoiceNumber = "INV-1001",
                Date = DateTime.Today,
                CustomerName = "John Doe",
                Items = new List<InvoiceItem>
                {
                    new InvoiceItem { Description = "Widget A", Quantity = 2, UnitPrice = 19.99m },
                    new InvoiceItem { Description = "Widget B", Quantity = 5, UnitPrice = 9.50m },
                    new InvoiceItem { Description = "Service C", Quantity = 1, UnitPrice = 150.00m }
                }
            };

            // Populate the template with the data using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object; the third parameter is optional
            // and can be used to reference the object by name inside the template.
            engine.BuildReport(doc, data, "invoice");

            // Configure SVG save options.
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                // Render text as placed glyphs so the output looks like an image.
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                // Optional: remove page borders and fit to viewport.
                ShowPageBorder = false,
                FitToViewPort = true
            };

            // Save the populated document as SVG.
            doc.Save("InvoiceOutput.svg", svgOptions);
        }
    }
}
