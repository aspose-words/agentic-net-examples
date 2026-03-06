using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReport
{
    // Simple POCO that represents a row in the in‑table list.
    public class Item
    {
        public string Name { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains the in‑table list tag (e.g. <<foreach [ds.Items]>><<[Name]>><</foreach>>).
            Document doc = new Document("Template.docx");

            // Prepare the data source – a list of Item objects.
            List<Item> items = new List<Item>
            {
                new Item { Name = "Alpha" },
                new Item { Name = "Beta" },
                new Item { Name = "Gamma" }
            };

            // Wrap the list in a container object so the template can reference it via a name (here "ds").
            var dataSource = new { Items = items };

            // Build the report using the ReportingEngine.
            // The third argument ("ds") is the name used inside the template to refer to the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "ds");

            // Save the populated document.
            doc.Save("Report.docx");
        }
    }
}
