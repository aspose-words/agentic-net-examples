using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data class that will be used as the data source for the template.
    public class InvoiceData
    {
        public string InvoiceNumber { get; set; }
        public string CustomerName { get; set; }
        public DateTime InvoiceDate { get; set; }
        public decimal TotalAmount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains reporting tags, e.g. <<[ds.InvoiceNumber]>>.
            string templatePath = @"C:\Templates\InvoiceTemplate.docx";

            // Path where the resulting PDF will be saved.
            string outputPdfPath = @"C:\Output\InvoiceReport.pdf";

            // Create a data object with sample values.
            InvoiceData data = new InvoiceData
            {
                InvoiceNumber = "INV-2023-001",
                CustomerName = "Acme Corp.",
                InvoiceDate = DateTime.Today,
                TotalAmount = 1234.56m
            };

            // Load the template document.
            Document doc = new Document(templatePath);

            // Populate the template using the ReportingEngine.
            // The third parameter ("ds") is the name used to reference the data source inside the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "ds");

            // Save the populated document as PDF.
            // Using SaveFormat.Pdf ensures the output format is PDF.
            doc.Save(outputPdfPath, SaveFormat.Pdf);
        }
    }
}
