using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class that will be used as the data source for the LINQ reporting engine.
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the PDF template that contains LINQ Reporting tags (e.g. <<[ds.Name]>>).
            //    The Document constructor handles both creation and loading.
            Document doc = new Document("Template.pdf");

            // 2. Prepare a sequential data source – a list of Item objects.
            List<Item> data = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 10 },
                new Item { Name = "Banana", Quantity = 20 },
                new Item { Name = "Cherry", Quantity = 30 }
            };

            // 3. Populate the template using the ReportingEngine.
            //    The data source name ("ds") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "ds");

            // 4. Save the populated document as Markdown.
            //    Configure MarkdownSaveOptions to specify the output format.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown   // Ensure the format is Markdown.
            };
            doc.Save("Report.md", saveOptions);
        }
    }
}
