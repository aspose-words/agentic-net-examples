using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsConditionalRowExample
{
    // Simple data model for the template.
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }

    public class ReportData
    {
        public List<Item> Items { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains a table with a conditional row.
            // The template should have a row like:
            // <<foreach [Items]>><<[Name]>> - <<[Quantity]>>\n<</foreach>>
            Document template = new Document("TemplateWithConditionalRow.docx");

            // Prepare the data source.
            var data = new ReportData
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apples",  Quantity = 10 },
                    new Item { Name = "Bananas", Quantity = 20 },
                    new Item { Name = "Cherries", Quantity = 30 }
                }
            };

            // Build the report – the ReportingEngine will repeat the table row for each item.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "ds"); // "ds" is the name used in the template (e.g., <<foreach [ds.Items]>>)

            // Save the populated document.
            template.Save("ReportWithConditionalRows.docx");
        }
    }
}
