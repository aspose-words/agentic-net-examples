using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model for the report.
    public class ReportData
    {
        // Collection that will be bound to a table in the DOTX template.
        public List<Item> Items { get; set; } = new List<Item>();
    }

    // Represents a single row in the table.
    public class Item
    {
        public string Name { get; set; }
        public decimal Price { get; set; }

        // This property can be used in the template for a conditional block.
        public bool IsExpensive => Price > 100m;
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare the data source.
            var data = new ReportData
            {
                Items = new List<Item>
                {
                    new Item { Name = "Basic Pen", Price = 2.5m },
                    new Item { Name = "Luxury Notebook", Price = 150m },
                    new Item { Name = "Standard Ruler", Price = 5m },
                    new Item { Name = "Premium Laptop", Price = 1250m }
                }
            };

            // 2. Load the DOTX template that contains the table and conditional block.
            // The template should have a LINQ Reporting block like:
            // <<foreach [in Items]>>
            //   <<[Name]>>
            //   <<[Price]:c>>
            //   <<if [IsExpensive]>><<[Price]:c>> (Expensive) <<endif>>
            // <<endforeach>>
            Document template = new Document("Template.dotx");

            // 3. Create the reporting engine and configure options if needed.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional blocks are omitted.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // 4. Build the report by populating the template with the data source.
            engine.BuildReport(template, data);

            // 5. Save the generated document.
            template.Save("ReportResult.docx");
        }
    }
}
