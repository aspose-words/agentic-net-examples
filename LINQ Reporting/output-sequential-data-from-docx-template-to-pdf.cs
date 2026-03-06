using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    // Simple POCO that will be used as the data source for the template.
    public class InvoiceData
    {
        public string CustomerName { get; set; }
        public string Address { get; set; }
        public DateTime InvoiceDate { get; set; }
        public decimal TotalAmount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags,
            // e.g. <<[Data.CustomerName]>>, <<[Data.Address]>>, etc.
            string templatePath = @"C:\Templates\InvoiceTemplate.docx";

            // Load the template document (lifecycle rule: use Document(string) constructor).
            Document doc = new Document(templatePath);

            // Prepare the data source object.
            InvoiceData data = new InvoiceData
            {
                CustomerName = "John Doe",
                Address = "123 Main St, Anytown",
                InvoiceDate = DateTime.Today,
                TotalAmount = 199.99m
            };

            // Populate the template with the data using the ReportingEngine (feature rule:
            // ReportingEngine.BuildReport(Document, object)).
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("Data") must match the name used in the template tags.
            engine.BuildReport(doc, data, "Data");

            // Save the populated document as PDF (lifecycle rule: use Document.Save(string, SaveFormat)).
            string outputPath = @"C:\Output\Invoice.pdf";
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
