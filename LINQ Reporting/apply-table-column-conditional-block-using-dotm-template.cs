using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTM template that contains a table with a conditional column.
        // Example of a conditional block inside a table cell in the template:
        // <<if [item.ShowQuantity]>><<[item.Quantity]>> <<endif>>
        Document template = new Document("Template.dotm");

        // Prepare a data source. Each Item has a flag (ShowQuantity) that determines
        // whether the quantity column should be displayed for that row.
        var data = new List<Item>
        {
            new Item { Name = "Apple",  Quantity = 10, ShowQuantity = true  },
            new Item { Name = "Banana", Quantity = 5,  ShowQuantity = false },
            new Item { Name = "Carrot", Quantity = 12, ShowQuantity = true  }
        };

        // Build the report using the ReportingEngine.
        // The third argument ("item") is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "item");

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Simple POCO class used as the data source for the report.
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public bool ShowQuantity { get; set; }
    }
}
