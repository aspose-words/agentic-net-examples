using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsPdfReport
{
    // Simple data class used as the data source for the template.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sequential data.
            var items = new List<Item>
            {
                new Item { Id = 1, Name = "First" },
                new Item { Id = 2, Name = "Second" },
                new Item { Id = 3, Name = "Third" }
            };

            // Load the PDF template (Aspose.Words can load PDF as a Document).
            Document template = new Document("Template.pdf");

            // Populate the template using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "items" can be referenced in the template as <<[items.Id]>> etc.
            engine.BuildReport(template, items, "items");

            // Save the populated document as PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            template.Save("Result.pdf", pdfOptions);
        }
    }
}
