using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace PdfTemplateToDocx
{
    // Sample data class that matches the fields used in the PDF template.
    public class InvoiceData
    {
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string CustomerName { get; set; }
        public List<LineItem> Items { get; set; }
        public decimal Total => CalculateTotal();

        private decimal CalculateTotal()
        {
            decimal sum = 0;
            if (Items != null)
            {
                foreach (var item in Items)
                    sum += item.Quantity * item.UnitPrice;
            }
            return sum;
        }
    }

    public class LineItem
    {
        public string Description { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains Aspose.Words reporting tags.
            string pdfTemplatePath = @"C:\Templates\InvoiceTemplate.pdf";

            // Load the PDF template as an Aspose.Words Document.
            Document doc = new Document(pdfTemplatePath);

            // Prepare the data source that will be merged into the template.
            InvoiceData data = new InvoiceData
            {
                InvoiceNumber = "INV-1001",
                InvoiceDate = DateTime.Today,
                CustomerName = "John Doe",
                Items = new List<LineItem>
                {
                    new LineItem { Description = "Widget A", Quantity = 2, UnitPrice = 19.99m },
                    new LineItem { Description = "Widget B", Quantity = 1, UnitPrice = 49.50m }
                }
            };

            // Use the ReportingEngine to populate the template with the data.
            ReportingEngine engine = new ReportingEngine();
            // The second overload allows referencing the data source by name ("data").
            engine.BuildReport(doc, data, "data");

            // Save the populated document as DOCX.
            string outputDocxPath = @"C:\Output\InvoiceResult.docx";
            doc.Save(outputDocxPath, SaveFormat.Docx);
        }
    }
}
