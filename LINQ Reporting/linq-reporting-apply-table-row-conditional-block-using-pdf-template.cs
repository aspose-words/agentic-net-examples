using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
    public class ReportData
    {
        // Collection that will be iterated in the template using a foreach block.
        public List<Item> Items { get; set; }
    }

    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }

        // This property is used by the template's conditional block (<<if [item.Show]>> ... <<endif>>).
        public bool Show => Quantity > 0;
    }

    class Program
    {
        static void Main()
        {
            // Load the Word template that contains a table with a conditional block.
            // The template should have tags like:
            // <<foreach [in Items]>>
            //   <<if [item.Show]>>
            //     <<[item.Name]>>
            //     <<[item.Quantity]>>
            //   <<endif>>
            // <<endforeach>>
            Document template = new Document("Template.docx");

            // Prepare the data source. LINQ is used here to demonstrate filtering,
            // but the conditional logic is also evaluated inside the template.
            ReportData data = new ReportData
            {
                Items = GetSampleItems()
                    .Where(i => i.Quantity >= 0) // Example LINQ filter; can be any condition.
                    .ToList()
            };

            // Configure the reporting engine. RemoveEmptyParagraphs ensures that rows
            // whose conditional block evaluates to false are removed cleanly.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report by populating the template with the data source.
            engine.BuildReport(template, data);

            // Save the populated document as PDF.
            template.Save("Report.pdf", SaveFormat.Pdf);
        }

        // Generates a sample collection of items.
        private static IEnumerable<Item> GetSampleItems()
        {
            return new List<Item>
            {
                new Item { Name = "Apples",  Quantity = 5 },
                new Item { Name = "Bananas", Quantity = 0 }, // Will be hidden by the conditional block.
                new Item { Name = "Cherries", Quantity = 12 },
                new Item { Name = "Dates", Quantity = -1 }   // Will be filtered out by LINQ.
            };
        }
    }
}
