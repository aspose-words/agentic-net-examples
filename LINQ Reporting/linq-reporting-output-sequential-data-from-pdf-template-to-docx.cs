using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace PdfToDocxReporting
{
    // Simple data class that will be used as the data source for the report.
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
            // Path to the PDF template that contains Aspose.Words reporting tags,
            // e.g. <<[ds.Description]>>, <<[ds.Total]>>, <<foreach [ds]>>...<<end>>
            const string pdfTemplatePath = @"C:\Templates\InvoiceTemplate.pdf";

            // Load the PDF template into an Aspose.Words Document.
            Document pdfTemplate = new Document(pdfTemplatePath);

            // Prepare a sequential list of data that will be merged into the template.
            List<InvoiceItem> items = new List<InvoiceItem>
            {
                new InvoiceItem { Description = "Widget A", Quantity = 3, UnitPrice = 19.99m },
                new InvoiceItem { Description = "Widget B", Quantity = 5, UnitPrice = 9.50m },
                new InvoiceItem { Description = "Service C", Quantity = 1, UnitPrice = 150.00m }
            };

            // Create an instance of the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove paragraphs that become empty after merging.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report by populating the PDF template with the data source.
            // The data source name "ds" is used inside the template tags.
            engine.BuildReport(pdfTemplate, items, "ds");

            // Save the populated document as DOCX.
            const string outputDocxPath = @"C:\Output\InvoiceReport.docx";
            pdfTemplate.Save(outputDocxPath, SaveFormat.Docx);
        }
    }
}
