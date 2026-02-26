using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains LINQ Reporting tags, e.g. <<[ds.TotalAmount]>>.
        string templatePath = "Template.docx";

        // Load the template document (lifecycle rule: load).
        Document doc = new Document(templatePath);

        // Prepare a data source for the report.
        var orders = new List<Order>
        {
            new Order { Id = 1, Amount = 120.5m, Date = new DateTime(2023, 1, 15) },
            new Order { Id = 2, Amount = 75.0m,  Date = new DateTime(2023, 2, 3) },
            new Order { Id = 3, Amount = 210.75m,Date = new DateTime(2023, 3, 22) }
        };

        // LINQ expression results that can be referenced from the template.
        var reportData = new
        {
            Orders = orders,
            TotalAmount = orders.Sum(o => o.Amount),
            OrderCount = orders.Count()
        };

        // Populate the template using the ReportingEngine (feature rule: BuildReport).
        ReportingEngine engine = new ReportingEngine();
        // The name "ds" must match the data source name used in the template tags.
        engine.BuildReport(doc, reportData, "ds");

        // Configure image save options to render the document as a multi‑page TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300 // optional: set DPI for higher quality.
            // No PageSet means all pages will be rendered.
        };

        // Save the populated document to TIFF (lifecycle rule: save).
        doc.Save("Report.tiff", saveOptions);
    }

    // Simple POCO representing an order; used in the data source.
    public class Order
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }
}
