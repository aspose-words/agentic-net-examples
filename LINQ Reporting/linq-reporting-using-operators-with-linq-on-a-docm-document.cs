using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple POCO that will be used as the data source for the report.
    public class ReportData
    {
        // The template will reference this property as <<[data.Items]>>
        public List<Item> Items { get; set; }
    }

    // Item class that represents a single record in the report.
    public class Item
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load a DOCM template that contains the reporting tags.
            // The template can contain a table with a foreach tag like:
            // <<foreach [data.Items]>><<[Category]>> <<[Name]>> <<[Price]:currency>><</foreach>>
            Document doc = new Document("Template.docm");

            // Prepare a raw collection of items.
            List<Item> allItems = new List<Item>
            {
                new Item { Category = "Fruit",   Name = "Apple",   Price = 1.20m },
                new Item { Category = "Fruit",   Name = "Banana",  Price = 0.80m },
                new Item { Category = "Veggie",  Name = "Carrot",  Price = 0.50m },
                new Item { Category = "Veggie",  Name = "Lettuce", Price = 1.00m },
                new Item { Category = "Fruit",   Name = "Orange",  Price = 1.50m },
                new Item { Category = "Snack",   Name = "Chips",   Price = 2.00m }
            };

            // Use LINQ operators to filter, sort and project the data.
            // Example: select only Fruit items, order by price descending.
            List<Item> filteredItems = allItems
                .Where(i => i.Category == "Fruit")          // filter
                .OrderByDescending(i => i.Price)           // sort
                .Select(i => new Item                       // projection (could be omitted here)
                {
                    Category = i.Category,
                    Name = i.Name,
                    Price = i.Price
                })
                .ToList();

            // Wrap the filtered collection in the ReportData object.
            ReportData dataSource = new ReportData
            {
                Items = filteredItems
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used inside the template to reference the data source.
            engine.BuildReport(doc, dataSource, "data");

            // Save the populated document. The output format is inferred from the extension.
            doc.Save("Report.docx");
        }
    }
}
