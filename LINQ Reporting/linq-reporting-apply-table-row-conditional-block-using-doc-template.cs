using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsConditionalTableRow
{
    // Data model used by the template.
    public class ReportData
    {
        public List<Item> Items { get; set; }
    }

    public class Item
    {
        public string Name { get; set; }
        public bool Show { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains a conditional block in a table row.
            // Example template syntax inside a table row:
            // <<if [ds.Items[i].Show]>><<[ds.Items[i].Name]>><<endif>>
            Document doc = new Document("Template.docx");

            // Prepare the data source.
            var data = new ReportData
            {
                Items = new List<Item>
                {
                    new Item { Name = "Item 1", Show = true },
                    new Item { Name = "Item 2", Show = false },
                    new Item { Name = "Item 3", Show = true }
                }
            };

            // Configure the reporting engine.
            var engine = new ReportingEngine
            {
                // Remove rows that become empty after the conditional block is evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name "ds" is used in the template.
            engine.BuildReport(doc, data, "ds");

            // Save the populated document.
            doc.Save("Result.docx");
        }
    }
}
